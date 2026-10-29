<#
===================================================================================================
  Build-Exe.ps1  -  turns SMS-Toolkit.ps1 into SMS-Toolkit.exe (Tecnica Systems icon)
---------------------------------------------------------------------------------------------------
  Run it after every change to SMS-Toolkit.ps1:
      right-click > Run with PowerShell
      or: powershell.exe -NoProfile -ExecutionPolicy Bypass -File Build-Exe.ps1

  What it does:
    1. checks SMS-Toolkit.ps1 for syntax errors (stops if there are any)
    2. installs the ps2exe module for the current user the first time (PowerShell Gallery)
    3. builds SMS-Toolkit.exe next to this script:
         - no console window, asks for administrator rights when started (UAC)
         - icon Assets\SMS-Toolkit.ico, version / company / product from SMS-Toolkit.ps1
    4. keeps the previous exe as SMS-Toolkit.previous.exe
===================================================================================================
#>
[CmdletBinding()]
param (
	[string]$Source,
	[string]$Output,
	[string]$Icon,
	[switch]$NoPause
)

$ErrorActionPreference = 'Stop'
# defaults next to this script ($PSScriptRoot is not reliable inside param() on Windows PowerShell 5.1)
$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Source) { $Source = Join-Path $here 'SMS-Toolkit.ps1' }
if (-not $Output) { $Output = Join-Path $here 'SMS-Toolkit.exe' }
if (-not $Icon) { $Icon = Join-Path $here 'Assets\SMS-Toolkit.ico' }

function Write_Step { param ([string]$Text) Write-Host ''; Write-Host ('==> ' + $Text) -ForegroundColor Cyan }
function Stop_Build
{
	param ([string]$Text)
	Write-Host ''
	Write-Host ('BUILD FAILED: ' + $Text) -ForegroundColor Red
	if (-not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
	exit 1
}

# ps2exe needs Windows PowerShell 5.1
if ($PSVersionTable.PSEdition -eq 'Core')
{
	$ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
	& $ps51 -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @PSBoundParameters
	exit $LASTEXITCODE
}

Write-Host 'SMS-TOOLKIT exe builder' -ForegroundColor Green

# ---------------------------------------------------------------------------------------------------
# 1. Source checks
# ---------------------------------------------------------------------------------------------------
Write_Step 'Checking the source'
if (-not (Test-Path -LiteralPath $Source)) { Stop_Build "source not found: $Source" }
if (-not (Test-Path -LiteralPath $Icon)) { Stop_Build "icon not found: $Icon" }
$tokens = $null; $errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($Source, [ref]$tokens, [ref]$errors)
if ($errors.Count)
{
	foreach ($e in ($errors | Select-Object -First 10)) { Write-Host ("  line {0}: {1}" -f $e.Extent.StartLineNumber, $e.Message) -ForegroundColor Red }
	Stop_Build "$($errors.Count) syntax error(s) in $(Split-Path $Source -Leaf)"
}
$text = [System.IO.File]::ReadAllText($Source)
$version = if ($text -match "(?m)^\s*Version\s*=\s*'(\d+(\.\d+){0,3})'") { $Matches[1] } else { '0.0.0' }
$parts = @($version.Split('.')); while ($parts.Count -lt 4) { $parts += '0' }
$fileVersion = ($parts | Select-Object -First 4) -join '.'
$company = if ($text -match "(?m)^\s*Company\s*=\s*'([^']+)'") { $Matches[1] } else { 'Tecnica Systems' }
$credits = if ($text -match "(?m)^\s*Credits\s*=\s*'([^']+)'") { ($Matches[1] -replace '\s+', ' ') } else { '' }
Write-Host ("  {0}  ({1:N0} lines, version {2})" -f (Split-Path $Source -Leaf), ($text -split "`n").Count, $fileVersion)

# ---------------------------------------------------------------------------------------------------
# 2. ps2exe module (current user, first run only)
# ---------------------------------------------------------------------------------------------------
Write_Step 'Checking the ps2exe module'
if (-not (Get-Module -ListAvailable -Name ps2exe))
{
	Write-Host '  ps2exe not installed - installing it for the current user from the PowerShell Gallery...'
	try
	{
		[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
		if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' }))
		{
			Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null
		}
		Install-Module -Name ps2exe -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
	}
	catch { Stop_Build "could not install ps2exe: $($_.Exception.Message)" }
}
Import-Module ps2exe
Write-Host ("  ps2exe {0}" -f (Get-Module ps2exe).Version)

# ---------------------------------------------------------------------------------------------------
# 3. Build
# ---------------------------------------------------------------------------------------------------
Write_Step 'Building the exe'
$previous = [System.IO.Path]::ChangeExtension($Output, '.previous.exe')
if (Test-Path -LiteralPath $Output)
{
	try { Copy-Item -LiteralPath $Output -Destination $previous -Force }
	catch { Stop_Build "the exe is in use - close SMS-TOOLKIT and run the build again ($($_.Exception.Message))" }
}
$build = @{
	inputFile    = $Source
	outputFile   = $Output
	iconFile     = $Icon
	noConsole    = $true      # GUI app: no console window
	noOutput     = $true      # stray output / errors never become message boxes
	noError      = $true
	requireAdmin = $true      # UAC prompt on start (the toolkit always runs elevated)
	STA          = $true      # WinForms
	title        = 'SMS-TOOLKIT'
	description  = 'SMS-TOOLKIT - Store Maintenance Toolkit'
	product      = 'SMS-TOOLKIT'
	company      = $company
	copyright    = ('{0} - {1}' -f $company, $credits)
	version      = $fileVersion
}
try { Invoke-ps2exe @build }
catch { Stop_Build $_.Exception.Message }
if (-not (Test-Path -LiteralPath $Output)) { Stop_Build 'ps2exe did not produce the exe' }

$exe = Get-Item -LiteralPath $Output
Write-Host ''
Write-Host ("BUILD OK: {0}  ({1:N0} KB, version {2}, built {3:yyyy-MM-dd HH:mm})" -f $exe.FullName, ($exe.Length / 1KB), $exe.VersionInfo.FileVersion, $exe.LastWriteTime) -ForegroundColor Green
if (Test-Path -LiteralPath $previous) { Write-Host ("  previous build kept as {0}" -f (Split-Path $previous -Leaf)) }
if (-not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
