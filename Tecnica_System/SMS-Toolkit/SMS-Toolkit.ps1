<#
===================================================================================================
  SMS-TOOLKIT  -  store maintenance toolkit (single portable file)
  Credits: Alex C.T  &  Enrique P.F
---------------------------------------------------------------------------------------------------
  Tools in this version:
    1. Database Maintenance - runs the store (STORESQL) or lane (LANESQL) maintenance / repair
       queries section by section, with a live terminal. The machine role (server or lane) and
       the database are detected automatically at startup.

  Run:   right-click > Run with PowerShell
         or: powershell.exe -NoProfile -ExecutionPolicy Bypass -File SMS-Toolkit.ps1
  Test:  powershell.exe -NoProfile -ExecutionPolicy Bypass -File SMS-Toolkit.ps1 -SelfTest
         (builds everything, runs detection, prints the result, no window)

  Needs Windows PowerShell 5.1 (WinForms + the built-in .NET SQL client). SQL is reached with
  Windows authentication as the user running the tool.
===================================================================================================
#>
[CmdletBinding()]
param (
	[switch]$SelfTest,
	[switch]$KeepConsole
)

# PowerShell 7 has no built-in System.Data.SqlClient: relaunch in Windows PowerShell 5.1
if ($PSVersionTable.PSEdition -eq 'Core')
{
	$ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
	$relaunch = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
	if ($SelfTest) { $relaunch += '-SelfTest' }
	if ($KeepConsole) { $relaunch += '-KeepConsole' }
	Start-Process -FilePath $ps51 -ArgumentList $relaunch
	return
}

# Always run elevated (rename, IP, registry, SQL sysadmin): relaunch with UAC, hidden console
if (-not $SelfTest)
{
	$isAdmin = $false
	try { $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) }
	catch { }
	if (-not $isAdmin)
	{
		try
		{
			if ($PSCommandPath)
			{
				# the .ps1: relaunch Windows PowerShell elevated, console hidden
				$ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
				$relaunch = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
				if ($KeepConsole) { $relaunch += '-KeepConsole' }
				Start-Process -FilePath $ps51 -ArgumentList $relaunch -Verb RunAs -WindowStyle $(if ($KeepConsole) { 'Normal' } else { 'Hidden' }) -ErrorAction Stop
			}
			else
			{
				# the .exe (Build-Exe.ps1 already asks for admin in its manifest - this is the fallback)
				Start-Process -FilePath ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Verb RunAs -ErrorAction Stop
			}
		}
		catch
		{
			Add-Type -AssemblyName System.Windows.Forms
			[void][System.Windows.Forms.MessageBox]::Show('SMS-TOOLKIT needs administrator rights to run.', 'Administrator required', 'OK', 'Warning')
		}
		return
	}
}

# ===================================================================================================
#                                      SECTION: Configuration
# ===================================================================================================

$script:App = @{
	Name    = 'SMS-TOOLKIT'
	Version = '0.5.0'
	Tagline = 'Store Maintenance Toolkit'
	Company = 'Tecnica Systems'
	Credits = 'Alex C.T  &  Enrique P.F'
}

# Packaged as an .exe (Build-Exe.ps1 / ps2exe, no console): Write-Host and Write-Warning would pop up
# message boxes there - send them to the terminal (or the file log before the window exists)
if ($Host.Name -match 'PS2EXE')
{
	function Write-Host
	{
		param ([Parameter(Position = 0, ValueFromRemainingArguments = $true)] $Object, $ForegroundColor, $BackgroundColor, [switch]$NoNewline, $Separator)
		$text = (@($Object) | ForEach-Object { [string]$_ }) -join ' '
		$kind = switch ([string]$ForegroundColor) { 'Green' { 'ok' } 'Red' { 'fail' } 'Yellow' { 'warn' } 'DarkYellow' { 'warn' } 'Cyan' { 'info' } 'Magenta' { 'info' } 'Gray' { 'muted' } 'DarkGray' { 'muted' } default { 'text' } }
		if ($script:UI.Term -and -not $script:UI.Term.IsDisposed) { Write_Term $text $kind } else { Write_LogFile $text }
	}
	function Write-Warning
	{
		param ([Parameter(Position = 0)] [string]$Message)
		if ($script:UI.Term -and -not $script:UI.Term.IsDisposed) { Write_Term $Message 'warn' } else { Write_LogFile ('WARNING: ' + $Message) }
	}
}

$script:Cfg = @{
	AlexRoot      = 'C:\Tecnica_Systems\Alex_C.T'
	StoremanRoots = @('C:\Storeman', 'D:\Storeman')
	# SQL login created by the maintenance script (used by the scheduled backup tasks)
	SqlUser       = 'Tecnica'
	SqlPass       = 'TB$upp0rT'
	ConnectSec    = 5
	# Title: logo navy letters; $true = thin white outline on the navy letters
	# (plain navy is hard to read on the dark background)
	TitleOutline  = $true
}

# Log + settings folder: the old tool's folder when possible, else %TEMP%
$script:Paths = @{ Log = $null; Setup = $null; Settings = $null }

# Dark theme
$script:T = $null   # filled by Initialize_Theme (needs System.Drawing)

# Runtime state
$script:Db = $null          # detection result (see Invoke_Detection)
$script:Steps = @()         # steps shown in the query list
$script:Mode = 'Maintenance'
$script:Run = $null         # active run: @{ PS; Handle; Sync; Started; Steps }
$script:Settings = @{ }
$script:UI = @{ }

# ---------------------------------------------------------------------------------------------------
# Tecnica Systems logo, embedded so the toolkit stays one file (decoded by Initialize_Branding)
# ---------------------------------------------------------------------------------------------------
# 128 px PNG, transparent (splash + title bar)
$script:LogoPngB64 = @'
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
YII=
'@
# ICO 16-64 px (window + taskbar icon)
$script:LogoIcoB64 = @'
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
'@

# ===================================================================================================
#                                    SECTION: Embedded SQL
# ---------------------------------------------------------------------------------------------------
#   Same SQL the old tool runs (Generate_SQL_Scripts: $ServerSQLScript and $LaneSQLFiltered).
#   Each "/* Name */" comment starts a section = one query in the list. Placeholders:
#   {{DB}} database, {{STORE}} 3-digit store number, {{SQLUSER}} / {{SQLPASS}} SQL login.
# ===================================================================================================

$script:SqlServerMaintenance = @'
/* Set memory configuration */
DECLARE @Memory25PercentMB BIGINT;
SELECT @Memory25PercentMB = (total_physical_memory_kb / 1024) * 25 / 100
FROM sys.dm_os_sys_memory;
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', @Memory25PercentMB;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Create "{{SQLUSER}}" user in the database */
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = '{{SQLUSER}}')
    CREATE LOGIN [{{SQLUSER}}] WITH PASSWORD = '{{SQLPASS}}', CHECK_POLICY = OFF;
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '{{SQLUSER}}')
    CREATE USER [{{SQLUSER}}] FOR LOGIN [{{SQLUSER}}];
IF IS_ROLEMEMBER('db_owner', '{{SQLUSER}}') = 0
    EXEC sp_addrolemember 'db_owner', '{{SQLUSER}}';

/* Create-Refresh ScaleCommApp Triggers */
IF OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.TBS_ITM_SMAppUPDATED (
        Id INT IDENTITY(1,1) PRIMARY KEY,
        CodeF01 VARCHAR(13) NULL,
        Sent BIT NOT NULL DEFAULT 0,
        SentAt DATETIME NOT NULL DEFAULT GETDATE()
    );
END

IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED') AND name = 'IDX_TBS_ITM_SMAppUPDATED_CodeF01')
    DROP INDEX IDX_TBS_ITM_SMAppUPDATED_CodeF01 ON dbo.TBS_ITM_SMAppUPDATED;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED') AND name = 'IDX_TBS_ITM_SMAppUPDATED_Sent')
    DROP INDEX IDX_TBS_ITM_SMAppUPDATED_Sent ON dbo.TBS_ITM_SMAppUPDATED;
IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID('dbo.TBS_ITM_SMAppUPDATED') AND name = 'IDX_TBS_ITM_SMAppUPDATED_SentAt')
    DROP INDEX IDX_TBS_ITM_SMAppUPDATED_SentAt ON dbo.TBS_ITM_SMAppUPDATED;

IF OBJECT_ID('dbo.SMApp_UpdateOBJ', 'TR') IS NOT NULL
    DROP TRIGGER dbo.SMApp_UpdateOBJ
GO
IF OBJECT_ID('dbo.SMApp_UpdatePOS', 'TR') IS NOT NULL
    DROP TRIGGER dbo.SMApp_UpdatePOS
GO
IF OBJECT_ID('dbo.SMApp_UpdatePrice', 'TR') IS NOT NULL
    DROP TRIGGER dbo.SMApp_UpdatePrice
GO
IF OBJECT_ID('dbo.SMApp_UpdateSCL_TXT', 'TR') IS NOT NULL
    DROP TRIGGER dbo.SMApp_UpdateSCL_TXT
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdateOBJ]
   ON  [dbo].[OBJ_TAB]
   AFTER INSERT,UPDATE
AS
BEGIN
	SET NOCOUNT ON;
	INSERT INTO dbo.TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
	SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000'
END
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdatePOS]
   ON  [dbo].[POS_TAB]
   AFTER INSERT,UPDATE
AS
BEGIN
	SET NOCOUNT ON;
	INSERT INTO dbo.TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
	SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000'
END
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdatePrice]
   ON  [dbo].[PRICE_TAB]
   AFTER INSERT,UPDATE
AS
BEGIN
	SET NOCOUNT ON;
	INSERT INTO dbo.TBS_ITM_SMAppUPDATED (CodeF01,Sent,SentAt)
	SELECT F01,0, GETDATE() FROM inserted WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=1 AND SUBSTRING(F01,9,5) = '00000'
END
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TRIGGER [dbo].[SMApp_UpdateSCL_TXT]
ON [dbo].[SCL_TXT_TAB]
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.TBS_ITM_SMAppUPDATED (CodeF01, Sent, SentAt)
    SELECT
        '002' + RIGHT('00000' + CONVERT(varchar(5), inserted.F267), 5) + '00000',
        0,
        GETDATE()
    FROM inserted
    JOIN dbo.OBJ_TAB
        ON '002' + RIGHT('00000' + CONVERT(varchar(5), inserted.F267), 5) + '00000' = OBJ_TAB.F01
    WHERE SUBSTRING(OBJ_TAB.F01,1,3) = '002'
      AND ISNUMERIC(SUBSTRING(OBJ_TAB.F01,4,5)) = 1
      AND SUBSTRING(OBJ_TAB.F01,9,5) = '00000';
END
GO

CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_CodeF01 ON dbo.TBS_ITM_SMAppUPDATED (CodeF01);
CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_Sent ON dbo.TBS_ITM_SMAppUPDATED (Sent);
CREATE INDEX IDX_TBS_ITM_SMAppUPDATED_SentAt ON dbo.TBS_ITM_SMAppUPDATED (SentAt);
GO

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Delete bad SMS items */
DELETE FROM OBJ_TAB WHERE F01='0020000000000';
DELETE FROM OBJ_TAB WHERE F01 LIKE '% %';
DELETE FROM OBJ_TAB WHERE LEN(F01)<>13;
DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000';
DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000';
DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);
DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);
DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);
DELETE FROM SCL_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB);
DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) <> '002';
DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000';
UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 ,SCL_TAB.F1001=1 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000'));
UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 ,SCL_TAB.F1001=1 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000'));
DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB);
DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB);
UPDATE SCL_TAB SET F267 = NULL, F1001 = 1 WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB);
UPDATE SCL_TAB SET F268 = NULL, F1001 = 1 WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB);
UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04, SCL_TXT_TAB.F1001=1 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0;
UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');
UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');
UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');
UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');
UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ');

/* Truncate PRICE_EVENT table for records older than 7 days */
IF OBJECT_ID('PRICE_EVENT','U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_EVENT','OBJECT','DELETE') = 1 DELETE FROM PRICE_EVENT WHERE F254 < DATEADD(DAY,-7,GETDATE());

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000);
DECLARE cmds CURSOR FOR
SELECT 'drop table [' + name + ']'
FROM sys.tables
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%'{{STORE_TABLES}}) AND DATEDIFF(DAY, create_date, GETDATE()) > 30;
OPEN cmds;
WHILE 1 = 1
BEGIN
    FETCH cmds INTO @cmd;
    IF @@fetch_status != 0 BREAK;
    EXEC(@cmd);
END;
CLOSE cmds;
DEALLOCATE cmds;

/* Rebuild indexes and update database statistics */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD';
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN';

/* Set DB to highest compatibility level supported by this SQL Server version */
DECLARE @ProductVersion varchar(128), @MajorVersion int, @CompatLevel int, @sql nvarchar(4000);
SET @ProductVersion = CONVERT(varchar(128), SERVERPROPERTY('ProductVersion'));
SET @MajorVersion = CONVERT(int, LEFT(@ProductVersion, CHARINDEX('.', @ProductVersion + '.') - 1));
SET @CompatLevel = CASE @MajorVersion WHEN 17 THEN 170 WHEN 16 THEN 160 WHEN 15 THEN 150 WHEN 14 THEN 140 WHEN 13 THEN 130 WHEN 12 THEN 120 WHEN 11 THEN 110 WHEN 10 THEN 100 WHEN 9 THEN 90 ELSE NULL END;
IF @CompatLevel IS NOT NULL
BEGIN
    SET @sql = N'ALTER DATABASE ' + QUOTENAME(N'{{DB}}') + N' SET COMPATIBILITY_LEVEL = ' + CONVERT(varchar(3), @CompatLevel);
    EXEC(@sql);
END

/* Shrink the main database file */
DBCC SHRINKFILE ({{DB}});

/* Shrink the database log file */
DBCC SHRINKFILE ({{DB}}_Log);

/* Restrict the indefinite log file growth */
ALTER DATABASE {{DB}} SET RECOVERY SIMPLE;
'@

$script:SqlLaneMaintenance = @'
/* Set memory configuration */
DECLARE @Memory25PercentMB BIGINT;
SELECT @Memory25PercentMB = (total_physical_memory_kb / 1024) * 25 / 100
FROM sys.dm_os_sys_memory;
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', @Memory25PercentMB;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Create "{{SQLUSER}}" user in the database */
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = '{{SQLUSER}}')
    CREATE LOGIN [{{SQLUSER}}] WITH PASSWORD = '{{SQLPASS}}', CHECK_POLICY = OFF;
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '{{SQLUSER}}')
    CREATE USER [{{SQLUSER}}] FOR LOGIN [{{SQLUSER}}];
IF IS_ROLEMEMBER('db_owner', '{{SQLUSER}}') = 0
    EXEC sp_addrolemember 'db_owner', '{{SQLUSER}}';

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Truncate PRICE_EVENT table for records older than 7 days */
IF OBJECT_ID('PRICE_EVENT','U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_EVENT','OBJECT','DELETE') = 1 DELETE FROM PRICE_EVENT WHERE F254 < DATEADD(DAY,-7,GETDATE());

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000)
DECLARE cmds CURSOR FOR
SELECT 'drop table [' + name + ']'
FROM sys.tables
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%'{{STORE_TABLES}}) AND DATEDIFF(DAY, create_date, GETDATE()) > 30
OPEN cmds
WHILE 1 = 1
BEGIN
FETCH cmds INTO @cmd
IF @@fetch_status != 0 BREAK
EXEC(@cmd)
END
CLOSE cmds;
DEALLOCATE cmds;

/* Cleaning HEADER_SAV */
IF OBJECT_ID('HEADER_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('HEADER_SAV', 'OBJECT', 'DELETE') = 1
    DELETE FROM HEADER_SAV
    WHERE (F903 = 'SVHOST' OR F903 = 'MPHOST' OR F903 = CONCAT('M', '{{STORE}}', '901'))
    AND (DATEDIFF(DAY, F907, GETDATE()) > 30 OR DATEDIFF(DAY, F909, GETDATE()) > 30);

/* Delete bad SMS items */
DELETE FROM OBJ_TAB WHERE F01='0020000000000'
DELETE FROM OBJ_TAB WHERE F01 LIKE '% %'
DELETE FROM OBJ_TAB WHERE LEN(F01)<>13
DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000'
DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000'
DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)
DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)
DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)
DELETE FROM SCL_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB)
DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) <> '002'
DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000'
UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000'))
UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000'))
DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB)
DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB)
UPDATE SCL_TAB SET F267 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB)
UPDATE SCL_TAB SET F268 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB)
UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0
UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')
UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' ')

/* Delete batches older than 7 days from the lane */
DELETE FROM Header_bat WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_dct WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_old WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_sav WHERE F909 < DATEADD(day, -7, GETDATE());

/* Rebuild indexes and update database statistics */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'

/* Shrink the main database file */
DBCC SHRINKFILE ({{DB}})

/* Shrink the database log file */
DBCC SHRINKFILE ({{DB}}_Log)

/* Restrict the indefinite log file growth */
ALTER DATABASE {{DB}} SET RECOVERY SIMPLE
'@

# Repair: guarded version of Repair_Database.sqi. Runs against master (the database itself may be
# unusable). Default = the safe steps; the data-loss repair is off and needs a typed confirmation.
#   Master      run while connected to master
#   SingleUser  puts the database in single-user mode -> the engine always returns it to
#               multi-user at the end, even after a failure or Stop
#   Danger      needs the typed confirmation; skipped when a selected backup failed
#   Backup      a failure here blocks the Danger steps
$script:RepairSteps = @(
	@{
		Name = 'Show who is using the database'; Default = $true; Master = $true
		Sql  = @'
SELECT s.session_id AS [Session], s.login_name AS [Login], s.host_name AS [Computer],
       s.program_name AS [Program], s.status AS [Status]
FROM sys.dm_exec_sessions s
WHERE s.database_id = DB_ID(N'{{DB}}') AND s.session_id <> @@SPID
ORDER BY s.session_id;
SELECT name AS [Database], state_desc AS [State], user_access_desc AS [Access], recovery_model_desc AS [Recovery]
FROM sys.databases WHERE name = N'{{DB}}';
'@
	},
	@{
		Name = 'Integrity check (read-only)'; Default = $true; Master = $true
		Sql  = @'
DBCC CHECKDB(N'{{DB}}') WITH NO_INFOMSGS, ALL_ERRORMSGS;
PRINT 'CHECKDB found no errors in {{DB}}.';
'@
	},
	@{
		Name = 'Back up the database (copy-only, before any repair)'; Default = $true; Master = $true; Backup = $true
		Sql  = @'
DECLARE @dir nvarchar(4000) = CONVERT(nvarchar(4000), SERVERPROPERTY('InstanceDefaultBackupPath'));
IF @dir IS NULL OR @dir = N''
    EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'BackupDirectory', @dir OUTPUT;
IF @dir IS NULL OR @dir = N'' SET @dir = N'C:\';
IF RIGHT(@dir, 1) <> N'\' SET @dir = @dir + N'\';
DECLARE @file nvarchar(4000) = @dir + N'{{DB}}_PreRepair_' + CONVERT(nvarchar(8), GETDATE(), 112) + N'_' + REPLACE(CONVERT(nvarchar(8), GETDATE(), 108), N':', N'') + N'.bak';
PRINT N'Backup file: ' + @file;
BACKUP DATABASE [{{DB}}] TO DISK = @file WITH COPY_ONLY, INIT, STATS = 10, CONTINUE_AFTER_ERROR;
'@
	},
	@{
		Name = 'Repair - rebuild (no data loss)'; Default = $true; Master = $true; SingleUser = $true
		Sql  = @'
ALTER DATABASE [{{DB}}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
BEGIN TRY
    DBCC CHECKDB(N'{{DB}}', REPAIR_REBUILD) WITH NO_INFOMSGS, ALL_ERRORMSGS;
END TRY
BEGIN CATCH
    DECLARE @m nvarchar(2048) = ERROR_MESSAGE();
    ALTER DATABASE [{{DB}}] SET MULTI_USER;
    RAISERROR(@m, 16, 1);
    RETURN;
END CATCH
ALTER DATABASE [{{DB}}] SET MULTI_USER;
PRINT 'REPAIR_REBUILD finished.';
'@
	},
	@{
		Name = 'Integrity check again (read-only)'; Default = $true; Master = $true
		Sql  = @'
DBCC CHECKDB(N'{{DB}}') WITH NO_INFOMSGS, ALL_ERRORMSGS;
PRINT 'CHECKDB found no errors in {{DB}}.';
'@
	},
	@{
		Name = 'Repair - ALLOW DATA LOSS (last resort)'; Default = $false; Master = $true; SingleUser = $true; Danger = $true
		Sql  = @'
ALTER DATABASE [{{DB}}] SET EMERGENCY;
ALTER DATABASE [{{DB}}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
BEGIN TRY
    DBCC CHECKDB(N'{{DB}}', REPAIR_ALLOW_DATA_LOSS) WITH NO_INFOMSGS, ALL_ERRORMSGS;
END TRY
BEGIN CATCH
    DECLARE @m nvarchar(2048) = ERROR_MESSAGE();
    BEGIN TRY ALTER DATABASE [{{DB}}] SET ONLINE; END TRY BEGIN CATCH END CATCH
    ALTER DATABASE [{{DB}}] SET MULTI_USER;
    RAISERROR(@m, 16, 1);
    RETURN;
END CATCH
IF DATABASEPROPERTYEX(N'{{DB}}', 'Status') <> 'ONLINE' ALTER DATABASE [{{DB}}] SET ONLINE;
ALTER DATABASE [{{DB}}] SET MULTI_USER;
PRINT 'REPAIR_ALLOW_DATA_LOSS finished - check the data in SMS.';
'@
	},
	@{
		Name = 'Return the database to normal use (online, multi-user)'; Default = $true; Master = $true
		Sql  = @'
IF DATABASEPROPERTYEX(N'{{DB}}', 'Status') <> 'ONLINE' ALTER DATABASE [{{DB}}] SET ONLINE;
IF DATABASEPROPERTYEX(N'{{DB}}', 'UserAccess') <> 'MULTI_USER' ALTER DATABASE [{{DB}}] SET MULTI_USER;
SELECT name AS [Database], state_desc AS [State], user_access_desc AS [Access] FROM sys.databases WHERE name = N'{{DB}}';
'@
	},
	@{
		Name = 'Shrink the database (optional)'; Default = $false; Master = $true
		Sql  = @'
DBCC SHRINKDATABASE(N'{{DB}}');
'@
	}
)

# ===================================================================================================
#                                  SECTION: Background scripts
# ---------------------------------------------------------------------------------------------------
#   Run in their own runspaces, so they must be self-contained (no functions of this file).
# ===================================================================================================

# ---------------------------------------------------------------------------------------------------
# Detection: find the storeman folder, decide server / lane, connect and test the database.
# Writes progress text to $Sync.Stage and returns one hashtable.
# ---------------------------------------------------------------------------------------------------
$script:DetectionScript = {
	param ($Sync, $Cfg)

	$r = @{
		Ok = $false; Role = $null; Server = $null; Database = $null; Store = $null; Terminal = $null
		Storeman = $null; Source = $null; Error = $null; Machine = $env:COMPUTERNAME
		SqlVersion = $null; Edition = $null; SizeMB = $null; State = $null; IsSysadmin = $null
		ConnectionString = $null; MasterConnectionString = $null; Notes = @()
	}

	$readIni = {
		param ([string]$Path)
		$h = @{ }
		try
		{
			foreach ($line in [System.IO.File]::ReadAllLines($Path))
			{
				$t = $line.Trim()
				if (-not $t -or $t.StartsWith(';') -or $t.StartsWith('[')) { continue }
				$i = $t.IndexOf('=')
				if ($i -lt 1) { continue }
				$k = $t.Substring(0, $i).Trim().ToUpperInvariant()
				if (-not $h.ContainsKey($k)) { $h[$k] = $t.Substring($i + 1).Trim() }
			}
		}
		catch { }
		$h
	}

	$connStr = {
		param ([string]$Server, [string]$Database)
		'Server={0};Database={1};Integrated Security=True;TrustServerCertificate=True;Connect Timeout={2};Application Name=SMS-Toolkit' -f $Server, $Database, $Cfg.ConnectSec
	}

	try
	{
		# ---- 1. storeman folder: share 'storeman' (registry, no slow SMB cmdlets) -> C:/D: -> drive scan
		$Sync.Stage = 'Locating the storeman folder...'
		$root = $null
		try
		{
			$key = Get-Item -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Shares' -ErrorAction Stop
			foreach ($n in $key.GetValueNames())
			{
				if ($n -ine 'storeman') { continue }
				foreach ($l in @($key.GetValue($n)))
				{
					if ($l -match '^Path=(.+)$' -and (Test-Path -LiteralPath (Join-Path $Matches[1] 'Startup.ini')))
					{
						$root = $Matches[1]; $r.Source = "share 'storeman'"
					}
				}
			}
		}
		catch { }
		if (-not $root)
		{
			foreach ($p in $Cfg.StoremanRoots)
			{
				if (Test-Path -LiteralPath (Join-Path $p 'Startup.ini')) { $root = $p; $r.Source = $p; break }
			}
		}
		if (-not $root)
		{
			foreach ($d in [System.IO.DriveInfo]::GetDrives())
			{
				if ($d.DriveType -ne 'Fixed' -or -not $d.IsReady) { continue }
				try
				{
					foreach ($dir in [System.IO.Directory]::GetDirectories($d.RootDirectory.FullName))
					{
						if ((Split-Path $dir -Leaf) -match '(?i)storeman' -and (Test-Path -LiteralPath (Join-Path $dir 'Startup.ini')))
						{
							$root = $dir; $r.Source = 'drive scan'; break
						}
					}
				}
				catch { }
				if ($root) { break }
			}
		}
		$r.Storeman = $root

		# ---- 2. role + database from Startup.ini (server: INFO_*901_SMSStart.ini first, like the old tool)
		$server = $null
		$db = $null
		if ($root)
		{
			$Sync.Stage = 'Reading Startup.ini...'
			$ini = & $readIni (Join-Path $root 'Startup.ini')
			if ($ini['STORE'] -match '^\d+$') { $r.Store = ([int]$ini['STORE']).ToString('000') }
			if ($ini['TER'] -match '^\d+$') { $r.Terminal = ([int]$ini['TER']).ToString('000') }
			$server = $ini['DBSERVER']
			$db = $ini['DBNAME']

			if ($r.Terminal) { $r.Role = if ([int]$r.Terminal -ge 900) { 'Server' } else { 'Lane' } }
			elseif ($db -match '(?i)^LANESQL$') { $r.Role = 'Lane' }
			elseif ($db) { $r.Role = 'Server' }

			if ($r.Role -eq 'Server')
			{
				$dbs = Join-Path $root 'Office\Dbs'
				$sms = $null
				if ($r.Store) { $sms = Join-Path $dbs ('INFO_{0}901_SMSStart.ini' -f $r.Store) }
				if (-not ($sms -and (Test-Path -LiteralPath $sms)))
				{
					$sms = Get-ChildItem -LiteralPath $dbs -Filter 'INFO_*901_SMSStart.ini' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
				}
				if ($sms)
				{
					$s = & $readIni $sms
					foreach ($k in $s.Keys)
					{
						if ($k -match 'SERVERNAME' -and $s[$k]) { $server = $s[$k] }
						if ($k -match 'DATABASENAME' -and $s[$k]) { $db = $s[$k] }
					}
					$r.Notes += "SQL details from $(Split-Path $sms -Leaf)"
				}
			}
			if (-not $db) { $db = if ($r.Role -eq 'Lane') { 'LANESQL' } else { 'STORESQL' } }
		}

		# (LOCAL) / localhost / empty = this machine; drop a stray trailing backslash
		if ($server)
		{
			$server = $server.Trim().TrimEnd('\')
			$server = $server -replace '^(?i)(\(local\)|localhost|127\.0\.0\.1)', '.'
		}
		if (-not $server) { $server = '.' }

		# ---- 3. connect; on failure (or no storeman) look for STORESQL / LANESQL on the local instances
		$tryConnect = {
			param ([string]$Srv, [string]$Db)
			$Sync.Stage = "Connecting to $Db on $Srv..."
			$cn = New-Object System.Data.SqlClient.SqlConnection (& $connStr $Srv 'master')
			try
			{
				$cn.Open()
				$cmd = $cn.CreateCommand()
				$cmd.CommandText = @"
SELECT CONVERT(nvarchar(128), SERVERPROPERTY('ProductVersion')), CONVERT(nvarchar(128), SERVERPROPERTY('Edition')),
       IS_SRVROLEMEMBER('sysadmin'), d.state_desc,
       (SELECT CAST(SUM(CAST(size AS bigint)) * 8 / 1024 AS int) FROM sys.master_files WHERE database_id = d.database_id)
FROM sys.databases d WHERE d.name = @db
"@
				$null = $cmd.Parameters.AddWithValue('@db', $Db)
				$rd = $cmd.ExecuteReader()
				try
				{
					if (-not $rd.Read()) { return "database $Db does not exist on $Srv" }
					$r.SqlVersion = [string]$rd.GetValue(0)
					$r.Edition = [string]$rd.GetValue(1)
					$r.IsSysadmin = ([string]$rd.GetValue(2) -eq '1')
					$r.State = [string]$rd.GetValue(3)
					$r.SizeMB = $rd.GetValue(4)
				}
				finally { $rd.Close() }
				$r.Server = $Srv
				$r.Database = $Db
				return $null
			}
			catch { return $_.Exception.InnerException.Message, $_.Exception.Message | Where-Object { $_ } | Select-Object -First 1 }
			finally { $cn.Dispose() }
		}

		$err = $null
		if ($root) { $err = & $tryConnect $server $db }
		if (-not $root -or $err)
		{
			if ($err) { $r.Notes += "Startup.ini target failed: $err" }
			$Sync.Stage = 'Searching local SQL Server instances...'
			$instances = @()
			foreach ($rk in @('HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'))
			{
				try { foreach ($n in (Get-Item -LiteralPath $rk -ErrorAction Stop).GetValueNames()) { $instances += if ($n -eq 'MSSQLSERVER') { '.' } else { ".\$n" } } }
				catch { }
			}
			$instances = @($instances | Select-Object -Unique)
			$candidates = if ($r.Role -eq 'Lane') { @('LANESQL', 'STORESQL') } else { @('STORESQL', 'LANESQL') }
			$found = $false
			foreach ($inst in $instances)
			{
				foreach ($cand in $candidates)
				{
					$e2 = & $tryConnect $inst $cand
					if (-not $e2)
					{
						$found = $true
						if (-not $r.Role) { $r.Role = if ($cand -eq 'LANESQL') { 'Lane' } else { 'Server' } }
						$r.Notes += "found $cand on local instance $inst"
						break
					}
				}
				if ($found) { break }
			}
			if (-not $found)
			{
				if (-not $instances) { $err = 'no SQL Server instance is installed on this machine' }
				elseif (-not $err) { $err = "STORESQL / LANESQL not found on: $($instances -join ', ')" }
				throw $err
			}
		}

		$r.ConnectionString = & $connStr $r.Server $r.Database
		$r.MasterConnectionString = & $connStr $r.Server 'master'
		$r.Ok = $true
	}
	catch { $r.Error = [string]$_.Exception.Message }
	$Sync.Stage = 'Done'
	$r
}

# ---------------------------------------------------------------------------------------------------
# Run engine: executes the selected steps in order and streams everything to $Sync.Queue.
# Each queue item: @{ T = step|sql|msg|table|ok|warn|fail|skip|info|done; X = text }.
# ---------------------------------------------------------------------------------------------------
$script:RunEngineScript = {
	param ($Sync, $TargetCs, $MasterCs, $Steps, $Secret, $RestoreSql)

	$global:Q = $Sync.Queue
	$global:Secret = $Secret
	$global:Rows = 0
	function Emit([string]$T, [string]$X)
	{
		if ($global:Secret) { $X = $X.Replace($global:Secret, '********') }
		$global:Q.Enqueue(@{ T = $T; X = $X })
	}
	function Fmt([TimeSpan]$D)
	{
		if ($D.TotalSeconds -lt 60) { return '{0:0.0} s' -f $D.TotalSeconds }
		return '{0}m {1:00}s' -f [int][Math]::Floor($D.TotalMinutes), $D.Seconds
	}

	$okCount = 0; $failCount = 0; $skipCount = 0
	$singleUserTouched = $false
	$blockDanger = $false
	$runStart = [datetime]::Now

	for ($i = 0; $i -lt $Steps.Count; $i++)
	{
		$s = $Steps[$i]
		$Sync.Index = $i
		$Sync.Current = $s.Name
		$Sync.StepStart = [datetime]::Now
		if ($Sync.Cancel)
		{
			Emit 'skip' ('Skipped (stopped): ' + $s.Name); $skipCount++; continue
		}
		if ($s.Danger -and $blockDanger)
		{
			Emit 'warn' ("Skipped: {0} - the backup failed, a data-loss repair without a backup is not allowed." -f $s.Name); $skipCount++; continue
		}

		Emit 'step' ('[{0}/{1}] {2}' -f ($i + 1), $Steps.Count, $s.Name)
		if ($Sync.ShowSql) { Emit 'sql' $s.Sql }
		if ($s.SingleUser) { $singleUserTouched = $true }

		$t0 = [datetime]::Now
		$global:Rows = 0
		$cn = New-Object System.Data.SqlClient.SqlConnection ($(if ($s.Master) { $MasterCs } else { $TargetCs }))
		$cn.add_InfoMessage({ param ($snd, $e) foreach ($er in $e.Errors) { $global:Q.Enqueue(@{ T = 'msg'; X = [string]$er.Message }) } })
		try
		{
			$cn.Open()
			foreach ($batch in [regex]::Split($s.Sql, '(?im)^\s*GO\s*;?\s*$'))
			{
				if ([string]::IsNullOrWhiteSpace($batch)) { continue }
				if ($Sync.Cancel) { throw 'Stopped by the user' }
				$cmd = $cn.CreateCommand()
				$cmd.CommandText = $batch
				$cmd.CommandTimeout = 0
				$cmd.add_StatementCompleted({ param ($snd, $e) $global:Rows += $e.RecordCount })
				$Sync.Cmd = $cmd
				$rd = $cmd.ExecuteReader()
				try
				{
					do
					{
						if ($rd.FieldCount -gt 0)
						{
							# result set -> small text table (max 50 rows)
							$cols = @(0..($rd.FieldCount - 1) | ForEach-Object { $rd.GetName($_) })
							$rows = New-Object System.Collections.ArrayList
							$more = 0
							while ($rd.Read())
							{
								if ($rows.Count -ge 50) { $more++; continue }
								$vals = @(0..($rd.FieldCount - 1) | ForEach-Object { $v = $rd.GetValue($_); if ($v -is [DBNull]) { '' } else { ([string]$v).Trim() } })
								[void]$rows.Add($vals)
							}
							if ($rows.Count -eq 0) { Emit 'table' ('  (no rows: ' + ($cols -join ', ') + ')') }
							else
							{
								$w = @(for ($c = 0; $c -lt $cols.Count; $c++) { [Math]::Min(40, [Math]::Max($cols[$c].Length, (@($rows | ForEach-Object { $_[$c].Length }) | Measure-Object -Maximum).Maximum)) })
								$fmtRow = { param ($vals) '  ' + ((0..($cols.Count - 1) | ForEach-Object { $t = [string]$vals[$_]; if ($t.Length -gt $w[$_]) { $t = $t.Substring(0, $w[$_] - 1) + '~' }; $t.PadRight($w[$_]) }) -join '  ') }
								$lines = @(& $fmtRow $cols)
								$lines += '  ' + ((0..($cols.Count - 1) | ForEach-Object { '-' * $w[$_] }) -join '  ')
								foreach ($row in $rows) { $lines += & $fmtRow $row }
								if ($more) { $lines += "  ... $more more row(s)" }
								Emit 'table' ($lines -join "`n")
							}
						}
					} while ($rd.NextResult())
				}
				finally { $rd.Close(); $Sync.Cmd = $null }
			}
			$rowsText = if ($global:Rows -gt 0) { ' - {0:N0} row(s) affected' -f $global:Rows } else { '' }
			Emit 'ok' ('{0} - done in {1}{2}' -f $s.Name, (Fmt ([datetime]::Now - $t0)), $rowsText)
			$okCount++
		}
		catch
		{
			$m = [string]$_.Exception.Message
			if ($_.Exception.InnerException) { $m = [string]$_.Exception.InnerException.Message }
			if ($Sync.Cancel)
			{
				Emit 'warn' ('{0} - stopped after {1}' -f $s.Name, (Fmt ([datetime]::Now - $t0))); $failCount++
			}
			else
			{
				Emit 'fail' ('{0} - FAILED after {1}: {2}' -f $s.Name, (Fmt ([datetime]::Now - $t0)), $m); $failCount++
				if ($s.Backup) { $blockDanger = $true }
			}
		}
		finally { $cn.Dispose() }
	}

	# Safety net: never leave the database in single-user / emergency mode
	if ($singleUserTouched -and $RestoreSql)
	{
		$cn = New-Object System.Data.SqlClient.SqlConnection $MasterCs
		try
		{
			$cn.Open()
			$cmd = $cn.CreateCommand(); $cmd.CommandText = $RestoreSql; $cmd.CommandTimeout = 120
			$null = $cmd.ExecuteNonQuery()
			Emit 'info' 'Checked: the database is online and in multi-user mode.'
		}
		catch { Emit 'fail' ('Could NOT return the database to multi-user mode: {0} - run "Return the database to normal use".' -f $_.Exception.Message) }
		finally { $cn.Dispose() }
	}

	$Sync.Index = $Steps.Count
	$kind = if ($failCount -gt 0) { 'fail' } elseif ($Sync.Cancel) { 'warn' } else { 'ok' }
	Emit 'done' ('{0}|Finished in {1}: {2} OK, {3} failed, {4} skipped' -f $kind, (Fmt ([datetime]::Now - $runStart)), $okCount, $failCount, $skipCount)
	$Sync.Done = $true
}

# ===================================================================================================
#                                     SECTION: Core functions
# ===================================================================================================

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Initialize_Paths - log + settings folder (old tool's folder, else %TEMP%\SMS-Toolkit)
# ---------------------------------------------------------------------------------------------------
function Initialize_Paths
{
	$base = $script:Cfg.AlexRoot
	try
	{
		foreach ($d in @((Join-Path $base 'Log'), (Join-Path $base 'Setup_Files')))
		{
			if (-not (Test-Path -LiteralPath $d)) { $null = New-Item -Path $d -ItemType Directory -Force -ErrorAction Stop }
		}
		$script:Paths.Log = Join-Path $base 'Log'
		$script:Paths.Setup = Join-Path $base 'Setup_Files'
	}
	catch
	{
		$tmp = Join-Path $env:TEMP 'SMS-Toolkit'
		$null = New-Item -Path $tmp -ItemType Directory -Force -ErrorAction SilentlyContinue
		$script:Paths.Log = $tmp
		$script:Paths.Setup = $tmp
	}
	$script:Paths.Settings = Join-Path $script:Paths.Setup 'SMS-Toolkit.settings.json'
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Write_LogFile - one line in Log\SMS-Toolkit_<date>.log (never throws)
# ---------------------------------------------------------------------------------------------------
function Write_LogFile
{
	param ([string]$Text)
	try
	{
		$path = Join-Path $script:Paths.Log ('SMS-Toolkit_{0:yyyy-MM-dd}.log' -f [datetime]::Now)
		$lines = foreach ($l in ($Text -split "`r?`n")) { '{0:HH:mm:ss}  {1}' -f [datetime]::Now, $l }
		[System.IO.File]::AppendAllText($path, (($lines -join "`r`n") + "`r`n"), [System.Text.Encoding]::UTF8)
	}
	catch { }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Protect_Text - masks the SQL login password in anything shown or logged
# ---------------------------------------------------------------------------------------------------
function Protect_Text
{
	param ([string]$Text)
	if ($script:Cfg.SqlPass) { return $Text.Replace($script:Cfg.SqlPass, '********') }
	$Text
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Load_Settings / Save_Settings - remembered query selections per role + mode
# ---------------------------------------------------------------------------------------------------
function Load_Settings
{
	$script:Settings = @{ }
	try
	{
		if (Test-Path -LiteralPath $script:Paths.Settings)
		{
			$obj = Get-Content -LiteralPath $script:Paths.Settings -Raw | ConvertFrom-Json
			foreach ($p in $obj.PSObject.Properties)
			{
				$inner = @{ }
				foreach ($q in $p.Value.PSObject.Properties) { $inner[$q.Name] = [bool]$q.Value }
				$script:Settings[$p.Name] = $inner
			}
		}
	}
	catch { $script:Settings = @{ } }
}

function Save_Settings
{
	try { $script:Settings | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $script:Paths.Settings -Encoding UTF8 }
	catch { }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Get_SqlSections - splits SQL text into steps at each /* Name */ comment
# ---------------------------------------------------------------------------------------------------
function Get_SqlSections
{
	param ([string]$Sql)
	$pattern = '(?s)/\*\s*(?<Name>[^*/]+?)\s*\*/\s*(?<Body>.*?)(?=/\*|$)'
	$list = @()
	foreach ($m in [regex]::Matches($Sql, $pattern))
	{
		$body = $m.Groups['Body'].Value.Trim()
		if (-not $body) { continue }
		$list += @{ Name = $m.Groups['Name'].Value.Trim(); Sql = $body; Default = $true; Master = $false }
	}
	$list
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Expand_Sql - fills the {{...}} placeholders for the detected database
# ---------------------------------------------------------------------------------------------------
function Expand_Sql
{
	param ([string]$Sql)
	$db = if ($script:Db -and $script:Db.Database) { $script:Db.Database } else { 'STORESQL' }
	$store = if ($script:Db) { [string]$script:Db.Store } else { '' }
	# Without a store number 'M%' / 'R%' would match far too many tables: leave those two out
	$storeTables = if ($store) { " OR name LIKE 'M$store%' OR name LIKE 'R$store%'" } else { '' }
	$Sql.Replace('{{DB}}', $db).Replace('{{STORE}}', $store).Replace('{{STORE_TABLES}}', $storeTables).
	Replace('{{SQLUSER}}', $script:Cfg.SqlUser).Replace('{{SQLPASS}}', $script:Cfg.SqlPass)
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Get_ModeSteps - the steps of a mode for the detected role
# ---------------------------------------------------------------------------------------------------
function Get_ModeSteps
{
	param ([string]$Mode, [string]$CustomFile)
	switch ($Mode)
	{
		'Maintenance'
		{
			$src = if ($script:Db -and $script:Db.Role -eq 'Lane') { $script:SqlLaneMaintenance } else { $script:SqlServerMaintenance }
			return @(Get_SqlSections (Expand_Sql $src))
		}
		'Repair'
		{
			return @($script:RepairSteps | ForEach-Object {
					$c = $_.Clone(); $c.Sql = Expand_Sql $c.Sql; $c
				})
		}
		'Custom'
		{
			$text = [System.IO.File]::ReadAllText($CustomFile)
			$steps = @(Get_SqlSections $text)
			if ($steps.Count -eq 0) { $steps = @(@{ Name = ('Entire file ({0})' -f (Split-Path $CustomFile -Leaf)); Sql = $text; Default = $true; Master = $false }) }
			foreach ($s in $steps)
			{
				# Same guards as Repair for anything that can lose data or lock everyone out
				if ($s.Sql -match '(?i)REPAIR_ALLOW_DATA_LOSS|DROP\s+DATABASE|SET\s+EMERGENCY') { $s.Danger = $true; $s.Default = $false }
				if ($s.Sql -match '(?i)SET\s+SINGLE_USER|SET\s+EMERGENCY') { $s.SingleUser = $true }
				if ($s.Sql -match '(?i)\bKILL\s+#') { $s.Default = $false; $s.Name += '  [placeholder #]' }
			}
			return $steps
		}
	}
}

# ===================================================================================================
#                     SECTION: Mini Ghost core (ported verbatim from Mini_Ghost.ps1)
# ---------------------------------------------------------------------------------------------------
#   Source: MiniGhost SCRIPT v1.3.8 (2026-04-15) by Alex_C.T. Business logic UNCHANGED - edit only in
#   consultation with Alex_C.T. These functions read the script-scope paths set by
#   Initialize_MiniGhost_Context and log with Write-Host (streamed to the terminal by the runner).
# ===================================================================================================

# ===================================================================================================
#                                FUNCTION: Set_Operation_Status
# ---------------------------------------------------------------------------------------------------
# Description:
# Central helper to update the shared operation-status table only when the requested key exists.
# ===================================================================================================

function Set_Operation_Status
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus,
		[Parameter(Mandatory = $true)]
		[string]$Key,
		[Parameter(Mandatory = $true)]
		[string]$Status,
		[Parameter(Mandatory = $false)]
		[string]$Message = "",
		[Parameter(Mandatory = $false)]
		[string]$Details = ""
	)
	
	if ($OperationStatus -and $OperationStatus.ContainsKey($Key))
	{
		$OperationStatus[$Key].Status = $Status
		$OperationStatus[$Key].Message = $Message
		$OperationStatus[$Key].Details = $Details
	}
}

# ===================================================================================================
#                                  FUNCTION: Set_Label_Text
# ---------------------------------------------------------------------------------------------------
# Description:
# Updates a WinForms label and optionally refreshes its parent layout.
# ===================================================================================================

function Set_Label_Text
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		$Label,
		[Parameter(Mandatory = $true)]
		[string]$Text,
		[Parameter(Mandatory = $false)]
		[switch]$RefreshParent,
		[Parameter(Mandatory = $false)]
		[switch]$ProcessEvents
	)
	
	if (-not $Label) { return }
	
	$Label.Text = $Text
	$Label.Refresh()
	
	if ($RefreshParent -and $Label.Parent)
	{
		$Label.Parent.PerformLayout()
		$Label.Parent.Refresh()
	}
	
	if ($ProcessEvents)
	{
		[System.Windows.Forms.Application]::DoEvents()
	}
}

# ===================================================================================================
#                              FUNCTION: Get_Effective_Machine_Name
# ---------------------------------------------------------------------------------------------------
# Description:
# Returns the pending/new machine name when one exists; otherwise returns the current computer name.
# ===================================================================================================

function Get_Effective_Machine_Name
{
	[CmdletBinding()]
	param ()
	
	if ($script:newMachineName -and -not [string]::IsNullOrWhiteSpace([string]$script:newMachineName))
	{
		return [string]$script:newMachineName
	}
	
	return $env:COMPUTERNAME
}

# ===================================================================================================
#                               FUNCTION: Get_Database_Connection_String
# ---------------------------------------------------------------------------------------------------
# Description:
#   Searches for the Startup.ini file in specified locations, extracts the DBNAME value,
#   constructs the connection string, and stores it in a script-level hashtable.
# ===================================================================================================

function Get_Database_Connection_String
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	if (-not $script:FunctionResults)
	{
		$script:FunctionResults = @{ }
	}
	
	$effectiveStartupIniPath = $StartupIniPath
	if ([string]::IsNullOrWhiteSpace($effectiveStartupIniPath))
	{
		$effectiveStartupIniPath = $startupIniPath
	}
	
	if (-not $effectiveStartupIniPath)
	{
		return
	}
	
	# Read the Startup.ini file
	try
	{
		$content = Get-Content -Path $effectiveStartupIniPath -ErrorAction Stop
		
		# Extract DBSERVER
		$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' }
		if ($dbServerLine)
		{
			$dbServer = $dbServerLine -replace '^DBSERVER=', ''
			$dbServer = $dbServer.Trim()
			if (-not $dbServer)
			{
				$dbServer = "localhost"
			}
		}
		else
		{
			$dbServer = "localhost"
		}
		
		# Extract DBNAME
		$dbNameLine = $content | Where-Object { $_ -match '^DBNAME=' }
		if ($dbNameLine)
		{
			$dbName = $dbNameLine -replace '^DBNAME=', ''
			$dbName = $dbName.Trim()
			if (-not $dbName)
			{
				return
			}
		}
		else
		{
			return
		}
	}
	catch
	{
		return
	}
	
	$configuredDbServer = $dbServer
	$resolvedRuntimeDbServer = $configuredDbServer
	$connectionString = "Server=$configuredDbServer;Database=$dbName;Integrated Security=True;Application Name=MiniGhost;"
	$candidateServers = New-Object System.Collections.Generic.List[string]
	$addCandidate = {
		param (
			[string]$Candidate
		)
		
		if ([string]::IsNullOrWhiteSpace($Candidate))
		{
			return
		}
		
		$normalizedCandidate = $Candidate.Trim()
		if ([string]::IsNullOrWhiteSpace($normalizedCandidate))
		{
			return
		}
		
		foreach ($existingCandidate in $candidateServers)
		{
			if ([string]::Equals($existingCandidate, $normalizedCandidate, [System.StringComparison]::OrdinalIgnoreCase))
			{
				return
			}
		}
		
		[void]$candidateServers.Add($normalizedCandidate)
	}
	
	$configuredHost = $configuredDbServer
	$instanceName = $null
	if ($configuredDbServer -match '\\')
	{
		$configuredParts = $configuredDbServer.Split('\', 2)
		$configuredHost = $configuredParts[0].Trim()
		if ($configuredParts.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($configuredParts[1]))
		{
			$instanceName = $configuredParts[1].Trim()
		}
	}
	
	$localAliases = @('localhost', '.', '(local)')
	if (-not [string]::IsNullOrWhiteSpace($env:COMPUTERNAME))
	{
		$localAliases += $env:COMPUTERNAME
	}
	
	$isConfiguredLocal = $false
	foreach ($alias in $localAliases)
	{
		if ([string]::Equals($configuredHost, $alias, [System.StringComparison]::OrdinalIgnoreCase))
		{
			$isConfiguredLocal = $true
			break
		}
	}
	
	if ($isConfiguredLocal)
	{
		& $addCandidate $configuredDbServer
	}
	
	if ($instanceName)
	{
		foreach ($alias in $localAliases)
		{
			& $addCandidate ("{0}\{1}" -f $alias, $instanceName)
		}
	}
	else
	{
		foreach ($alias in $localAliases)
		{
			& $addCandidate $alias
		}
		
		foreach ($registryPath in @(
				'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL',
				'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'
			))
		{
			if (-not (Test-Path $registryPath)) { continue }
			
			try
			{
				$itemProps = Get-ItemProperty -Path $registryPath -ErrorAction Stop
				foreach ($prop in $itemProps.PSObject.Properties)
				{
					if ($prop.Name -in @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider'))
					{
						continue
					}
					
					if ([string]::IsNullOrWhiteSpace($prop.Name))
					{
						continue
					}
					
					$instanceCandidate = $prop.Name.Trim()
					foreach ($alias in $localAliases)
					{
						& $addCandidate ("{0}\{1}" -f $alias, $instanceCandidate)
					}
				}
			}
			catch { }
		}
	}
	
	& $addCandidate $configuredDbServer
	
	foreach ($candidateServer in $candidateServers)
	{
		$testConnection = $null
		try
		{
			$testConnectionString = "Server=$candidateServer;Database=$dbName;Integrated Security=True;Connect Timeout=3;Application Name=MiniGhost;"
			$testConnection = New-Object System.Data.SqlClient.SqlConnection($testConnectionString)
			$testConnection.Open()
			
			$resolvedRuntimeDbServer = $candidateServer
			$connectionString = "Server=$candidateServer;Database=$dbName;Integrated Security=True;Application Name=MiniGhost;"
			break
		}
		catch { }
		finally
		{
			if ($testConnection)
			{
				$testConnection.Close()
				$testConnection.Dispose()
			}
		}
	}
	
	# Store configured and runtime database context in the FunctionResults hashtable
	$script:FunctionResults['ConfiguredDBSERVER'] = $configuredDbServer
	$script:FunctionResults['DBSERVER'] = $configuredDbServer
	$script:FunctionResults['RuntimeDBSERVER'] = $resolvedRuntimeDbServer
	$script:FunctionResults['DBNAME'] = $dbName
	$script:FunctionResults['ConnectionString'] = $connectionString
}

# ===================================================================================================
#                          FUNCTION: Ensure_Database_Connection_Context
# ---------------------------------------------------------------------------------------------------
# Description:
# Lazily refreshes DBSERVER / DBNAME / ConnectionString from Startup.ini so SQL-specific work can
# initialize only when needed instead of front-loading that work during GUI startup.
# ===================================================================================================

function Ensure_Database_Connection_Context
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	if (-not $script:FunctionResults)
	{
		$script:FunctionResults = @{ }
	}
	
	$connectionStringValue = $null
	if ($script:FunctionResults.ContainsKey('ConnectionString') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConnectionString']))
	{
		$connectionStringValue = [string]$script:FunctionResults['ConnectionString']
	}
	
	$dbNameValue = $null
	if ($script:FunctionResults.ContainsKey('DBNAME') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBNAME']))
	{
		$dbNameValue = [string]$script:FunctionResults['DBNAME']
	}
	
	$dbServerValue = $null
	if ($script:FunctionResults.ContainsKey('DBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBSERVER']))
	{
		$dbServerValue = [string]$script:FunctionResults['DBSERVER']
	}
	
	if ([string]::IsNullOrWhiteSpace($connectionStringValue) -or [string]::IsNullOrWhiteSpace($dbNameValue) -or [string]::IsNullOrWhiteSpace($dbServerValue))
	{
		Get_Database_Connection_String -StartupIniPath $StartupIniPath
		
		if ($script:FunctionResults.ContainsKey('ConnectionString') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConnectionString']))
		{
			$connectionStringValue = [string]$script:FunctionResults['ConnectionString']
		}
	}
	
	$script:connectionString = $connectionStringValue
	return $connectionStringValue
}

# ===================================================================================================
#                                       FUNCTION: Get_Store_Number_From_INI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store number from the startup.ini file.
# ===================================================================================================

function Get_Store_Number_From_INI
{
	[CmdletBinding()]
	param (
		[switch]$UpdateLabel
	)
	
	# Ensure FunctionResults exists (PS 5.1 safe)
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	
	# Default/fallback
	$script:FunctionResults['StoreNumber'] = "N/A"
	$foundStore = $null
	
	if (Test-Path $startupIniPath)
	{
		try
		{
			$iniContent = Get-Content $startupIniPath -ErrorAction Stop
			
			foreach ($line in $iniContent)
			{
				if ($line -match '^\s*STORE\s*=\s*(\d{3,4})\s*$')
				{
					$foundStore = $matches[1]
					$script:FunctionResults['StoreNumber'] = $foundStore
					break
				}
			}
		}
		catch
		{
			Write-Warning ("Failed to read startup.ini: {0}" -f $_)
		}
	}
	
	# Update label ONLY if requested
	if ($UpdateLabel -and (-not $SilentMode) -and $script:storeNumberLabel)
	{
		Set_Label_Text -Label $script:storeNumberLabel -Text "Store Number: $($script:FunctionResults['StoreNumber'])" -RefreshParent -ProcessEvents
	}
	
	if ($foundStore) { return $foundStore }
	return $null
}

# ===================================================================================================
#                                      FUNCTION: Get_Store_Name_From_INI
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the store name from the system.ini file.
#   Stores the result in $script:FunctionResults['StoreName'].
# ===================================================================================================

function Get_Store_Name_From_INI
{
	param (
		[string]$INIPath = $SystemIniPath
	)
	
	# Initialize StoreName
	$script:FunctionResults['StoreName'] = "N/A"
	
	if (Test-Path $INIPath)
	{
		$storeName = Select-String -Path $INIPath -Pattern "^NAME=" | ForEach-Object {
			$_.Line.Split('=')[1].Trim()
		}
		if ($storeName)
		{
			$script:FunctionResults['StoreName'] = $storeName
			# Write-Log "Store name found in system.ini: $storeName" "green"
		}
		else
		{
			# Write-Log "Store name not found in system.ini." "yellow"
		}
	}
	else
	{
		# Write-Log "INI file not found: $INIPath" "yellow"
	}
	
	# Update the storeNameLabel in the GUI
	if (-not $SilentMode -and $storeNameLabel -ne $null)
	{
		$storeNameLabel.Text = "Store Name: $($script:FunctionResults['StoreName'])"
		$form.Refresh()
		[System.Windows.Forms.Application]::DoEvents()
	}
}

# ===================================================================================================
#                                       FUNCTION: Get_Active_IP_Config
# ---------------------------------------------------------------------------------------------------
# Description:
#   Retrieves the active IP configuration for network adapters that are up and have valid IPv4 addresses.
#   Optimized for performance by prefiltering and minimizing pipeline overhead.
#   Attempts WMI first for compatibility, falls back to NetAdapter methods if WMI fails.
# ===================================================================================================

function Get_Active_IP_Config
{
	try
	{
		# WMI attempt: Get default route from Win32_IP4RouteTable
		$defaultRoute = Get-WmiObject -Class Win32_IP4RouteTable -ErrorAction Stop |
		Where-Object { $_.Destination -eq '0.0.0.0' -and $_.Mask -eq '0.0.0.0' -and $_.NextHop -ne '0.0.0.0' } |
		Sort-Object -Property Metric1 |
		Select-Object -First 1
		
		if ($defaultRoute -and $defaultRoute.InterfaceIndex)
		{
			$idx = [int]$defaultRoute.InterfaceIndex
			
			# Check adapter status via Win32_NetworkAdapter (NetConnectionStatus 2 = Connected)
			$ad = Get-WmiObject -Class Win32_NetworkAdapter -Filter "DeviceID = $idx" -ErrorAction SilentlyContinue
			if ($ad -and $ad.NetConnectionStatus -eq 2)
			{
				# Get config via Win32_NetworkAdapterConfiguration
				$wmiCfg = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "Index = $idx" -ErrorAction SilentlyContinue
				if ($wmiCfg -and $wmiCfg.IPEnabled -and $wmiCfg.DefaultIPGateway)
				{
					# Filter for valid IPv4 (non-APIPA)
					$validIPs = $wmiCfg.IPAddress | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' -and $_ -notmatch '^169\.254\.' }
					if ($validIPs)
					{
						# Construct a custom object similar to Get-NetIPConfiguration for consistency
						$ipv4Address = [PSCustomObject]@{
							IPAddress    = $validIPs[0]
							PrefixLength = ($wmiCfg.IPSubnet | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' })[0] | ForEach-Object {
								[int]([math]::Log([uint32]::MaxValue -bxor ([ipaddress]$_).Address, 2)) + 1
							}
						}
						$cfg = [PSCustomObject]@{
							InterfaceAlias	   = $ad.NetConnectionID
							InterfaceIndex	   = $idx
							IPv4Address	       = $ipv4Address
							IPv4DefaultGateway = [PSCustomObject]@{ NextHop = $wmiCfg.DefaultIPGateway[0] }
							DNSServer		   = $wmiCfg.DNSServerSearchOrder
							# Add more properties if needed for parity
						}
						return @($cfg) # Return as array for consistent .Count / [0]
					}
				}
			}
		}
	}
	catch
	{
		# Ignore and fall back to original methods
	}
	
	# =========================================================================================
	# FALLBACK: Original approach (NetRoute/NetAdapter/NetIPConfiguration)
	# =========================================================================================
	
	try
	{
		$defaultRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 -ErrorAction Stop |
		Where-Object { $_.NextHop -and $_.NextHop -ne '0.0.0.0' } |
		Sort-Object -Property RouteMetric, InterfaceMetric |
		Select-Object -First 1
		
		if ($defaultRoute -and $defaultRoute.InterfaceIndex)
		{
			$idx = [int]$defaultRoute.InterfaceIndex
			
			# Ensure adapter is Up (do NOT restrict to -Physical here; active interface might be Wi-Fi/VPN/etc.)
			$ad = Get-NetAdapter -InterfaceIndex $idx -ErrorAction SilentlyContinue
			if ($ad -and $ad.Status -eq 'Up')
			{
				# Make sure there is at least one valid IPv4 (non-APIPA) on this interface
				$ipRow = Get-NetIPAddress -InterfaceIndex $idx -AddressFamily IPv4 -ErrorAction SilentlyContinue |
				Where-Object { $_.IPAddress -and $_.IPAddress -notlike '169.254.*' -and $_.IPAddress -ne '0.0.0.0' } |
				Select-Object -First 1
				
				if ($ipRow)
				{
					$cfg = Get-NetIPConfiguration -InterfaceIndex $idx -ErrorAction SilentlyContinue
					if ($cfg -and $cfg.IPv4Address)
					{
						# Validate IPv4Address (can be array). If any non-APIPA exists, return it.
						$hasValid = $false
						foreach ($addr in @($cfg.IPv4Address))
						{
							if ($addr -and $addr.IPAddress -and ($addr.IPAddress -notlike '169.254.*'))
							{
								$hasValid = $true
								break
							}
						}
						
						if ($hasValid)
						{
							return @($cfg) # return as array for consistent .Count / [0]
						}
					}
				}
			}
		}
	}
	catch
	{
		# ignore and fall back
	}
	
	# =========================================================================================
	# SECONDARY FALLBACK: Physical + Up adapters -> Get-NetIPConfiguration
	# =========================================================================================
	
	# Prefilter by adapter status to reduce objects early
	$adapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object Status -eq 'Up'
	if (-not $adapters) { return $null }
	
	# Get only configs for up adapters
	$adapterNames = $adapters | Select-Object -ExpandProperty Name
	$ipConfigs = Get-NetIPConfiguration -InterfaceAlias $adapterNames -ErrorAction SilentlyContinue
	
	# Filter for valid IPv4 (not APIPA/169.254.x.x, not null)
	$validConfigs = $ipConfigs | Where-Object {
		$_.IPv4Address -and
		$_.IPv4Address.IPAddress -and
		($_.IPv4Address.IPAddress -notlike '169.254*')
	}
	
	if ($validConfigs) { return @($validConfigs) }
	return $null
}

# ===================================================================================================
# FUNCTION: Remove_Old_XZ_Folders
# ---------------------------------------------------------------------------------------------------
# Description:
# Removes old X* and Z* folders under \storeman\office that do NOT match the NEW store and allowed
# terminal numbers.
#
# Strict folder name pattern acted upon:
#   ^(X|Z)[A-Za-z](store 3-4 digits)(terminal 3 digits)$
# Examples: XF0231006, XW0242901, ZF1234006, ZW1234901
#
# Keeps ONLY folders that match:
#   - Store = new store (3-digit OR 4-digit representation)
#   - Terminal IN: current lane terminal (from MachineName), 900, 901
#
# Safeguard: refuses to run if extracted terminal is 901 unless -AllowBackoffice is passed.
# Automatic: no prompts.
# ===================================================================================================

function Remove_Old_XZ_Folders
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus,
		[Parameter(Mandatory = $false)]
		[switch]$AllowBackoffice
	)
	
	# -----------------------------
	# Init / status bucket
	# -----------------------------
	$deletedFolders = @()
	$failedToDeleteFolders = @()
	$anyFoldersSeen = $false
	$anyCandidatesMatched = $false
	
	if ($OperationStatus)
	{
		if (-not $OperationStatus.ContainsKey("OldXFoldersDeletion"))
		{
			$OperationStatus["OldXFoldersDeletion"] = [pscustomobject]@{
				Status  = ""
				Message = ""
				Details = ""
			}
		}
	}
	
	# -----------------------------
	# Validate store number
	# -----------------------------
	$storeNumberTrim = $StoreNumber
	if ($null -eq $storeNumberTrim) { $storeNumberTrim = "" }
	$storeNumberTrim = $storeNumberTrim.Trim()
	
	if ($storeNumberTrim -notmatch '^(?!0+$)\d{3,4}$')
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Invalid StoreNumber '$StoreNumber' (must be 3 or 4 digits, not all zeros)."
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: Invalid StoreNumber '$StoreNumber'." -ForegroundColor Red
		return
	}
	
	$storeInt = [int]$storeNumberTrim
	$store3 = $storeInt.ToString("D3")
	$store4 = $storeInt.ToString("D4")
	
	# -----------------------------
	# Pick base path
	# -----------------------------
	$possibleBasePaths = @("\\localhost\storeman\office", "C:\storeman\office", "D:\storeman\office")
	$basePath = $null
	foreach ($p in $possibleBasePaths)
	{
		if (Test-Path $p) { $basePath = $p; break }
	}
	
	if (-not $basePath)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "None of the base paths exist: $($possibleBasePaths -join ', ')"
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: None of the base paths exist: $($possibleBasePaths -join ', ')" -ForegroundColor Red
		return
	}
	
	# -----------------------------
	# Extract machine terminal (last 1-3 digits, padded to 3)
	# -----------------------------
	$mn = $MachineName
	if ($null -eq $mn) { $mn = "" }
	$mn = $mn.Trim()
	
	$mn = $mn -replace '^[\\\/]+', '' # strip leading \\ or /
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] } # keep host portion
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] } # drop domain
	$mn = $mn.Trim().ToUpper()
	
	if ($mn -notmatch '(\d{1,3})$')
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "MachineName '$MachineName' does not end with 1-3 digits; cannot extract terminal number."
			$OperationStatus["OldXFoldersDeletion"].Details = "Cannot proceed with folder deletion."
		}
		Write-Host "Failed: MachineName '$MachineName' cannot provide terminal digits." -ForegroundColor Red
		return
	}
	
	$machineNumber = ([int]$Matches[1]).ToString("D3")
	
	# Safeguard for backoffice unless explicitly allowed
	if ($machineNumber -eq "901" -and -not $AllowBackoffice)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Refusing to run because extracted terminal number is 901 (backoffice). Use -AllowBackoffice to override."
			$OperationStatus["OldXFoldersDeletion"].Details = "This could delete many lane folders. Override only if intentional."
		}
		Write-Host "Failed: Refusing to run on terminal 901 (backoffice safeguard). Use -AllowBackoffice to override." -ForegroundColor Red
		return
	}
	
	# Keep set: current lane + 900 + 901
	$keepStores = @($store3, $store4)
	$keepTerminals = @($machineNumber, "900", "901")
	
	# -----------------------------
	# Enumerate X* and Z* folders
	# -----------------------------
	$folders = @()
	
	$xFolders = Get-ChildItem -Path $basePath -Directory -Filter "X*" -ErrorAction SilentlyContinue
	if ($xFolders) { $folders += $xFolders }
	
	$zFolders = Get-ChildItem -Path $basePath -Directory -Filter "Z*" -ErrorAction SilentlyContinue
	if ($zFolders) { $folders += $zFolders }
	
	if (-not $folders -or $folders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Folders Found"
			$OperationStatus["OldXFoldersDeletion"].Message = "No X* or Z* folders were found under '$basePath'."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: No X* or Z* folders found under '$basePath'." -ForegroundColor Cyan
		return
	}
	
	$anyFoldersSeen = $true
	
	foreach ($folder in $folders)
	{
		$folderName = $folder.Name
		
		# STRICT pattern only: (X|Z)(letter)(store 3-4 digits)(terminal 3 digits)
		if ($folderName -match '^(?<prefix>[XZ][A-Za-z])(?<store>\d{3,4})(?<terminal>\d{3})$')
		{
			$folderStore = $matches['store']
			$folderTerminal = $matches['terminal']
			
			$storeOk = ($keepStores -contains $folderStore)
			$terminalOk = ($keepTerminals -contains $folderTerminal)
			
			# Delete if it does NOT match new store AND allowed terminal(s)
			if (-not ($storeOk -and $terminalOk))
			{
				$anyCandidatesMatched = $true
				
				$maxRetries = 3
				$retryCount = 0
				$deleted = $false
				
				while ($retryCount -lt $maxRetries -and -not $deleted)
				{
					try
					{
						Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
						$deletedFolders += $folderName
						$deleted = $true
					}
					catch
					{
						$retryCount++
						Start-Sleep -Seconds 2
						
						if ($retryCount -ge $maxRetries)
						{
							$failedToDeleteFolders += $folderName
							Write-Host "Failed to delete folder: $folderName. Error: $($_.Exception.Message)" -ForegroundColor Red
						}
					}
				}
			}
		}
	}
	
	# -----------------------------
	# Results
	# -----------------------------
	$resultMessage = ""
	if ($deletedFolders.Count -gt 0)
	{
		$resultMessage += "Deleted folders:`n$($deletedFolders -join "`n")`n"
	}
	if ($failedToDeleteFolders.Count -gt 0)
	{
		$resultMessage += "Failed to delete folders:`n$($failedToDeleteFolders -join "`n")`n"
	}
	
	if (-not $anyCandidatesMatched)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "No Matching Folders"
			$OperationStatus["OldXFoldersDeletion"].Message = "No X*/Z* folders needed deletion. Kept stores: $($keepStores -join ', ') terminals: $($keepTerminals -join ', ')."
			$OperationStatus["OldXFoldersDeletion"].Details = "Nothing to delete."
		}
		Write-Host "Info: Nothing to delete. Kept stores: $($keepStores -join ', ') terminals: $($keepTerminals -join ', ')." -ForegroundColor Cyan
		return
	}
	
	if ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -eq 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Successful"
			$OperationStatus["OldXFoldersDeletion"].Message = "Old X*/Z* folders deleted successfully."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Success: Old X*/Z* folders deleted successfully." -ForegroundColor Green
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Green }
	}
	elseif ($deletedFolders.Count -gt 0 -and $failedToDeleteFolders.Count -gt 0)
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Partial Failure"
			$OperationStatus["OldXFoldersDeletion"].Message = "Some old X*/Z* folders could not be deleted."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Warning: Some old X*/Z* folders could not be deleted." -ForegroundColor Yellow
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Yellow }
	}
	else
	{
		if ($OperationStatus)
		{
			$OperationStatus["OldXFoldersDeletion"].Status = "Failed"
			$OperationStatus["OldXFoldersDeletion"].Message = "Failed to delete any old X*/Z* folders."
			$OperationStatus["OldXFoldersDeletion"].Details = $resultMessage
		}
		Write-Host "Error: Failed to delete any old X*/Z* folders." -ForegroundColor Red
		if ($resultMessage) { Write-Host $resultMessage -ForegroundColor Red }
	}
}

# ===================================================================================================
#                                       FUNCTION: Execute_SQL_Commands
# ---------------------------------------------------------------------------------------------------
# Description:
#   Executes a given SQL command using the provided connection string.
# ===================================================================================================

function Execute_SQL_Commands
{
	param (
		[string]$commandText
	)
	
	$script:LastSqlExecutionError = $null
	$resolvedConnectionString = Ensure_Database_Connection_Context
	if ([string]::IsNullOrWhiteSpace($resolvedConnectionString))
	{
		$script:LastSqlExecutionError = "Database connection string is not available."
		return $false
	}
	
	$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$sqlConnection.ConnectionString = $resolvedConnectionString
	$sqlCommand = $sqlConnection.CreateCommand()
	$sqlCommand.CommandText = $commandText
	
	try
	{
		$sqlConnection.Open()
		$sqlCommand.ExecuteNonQuery() | Out-Null
		return $true # Indicate success
	}
	catch
	{
		$script:LastSqlExecutionError = $_.Exception.Message
		return $false # Indicate failure
	}
	finally
	{
		$sqlConnection.Close()
	}
}

# ===================================================================================================
#                             FUNCTION: Update_SQL_Tables_For_Store_Number_Change
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates STD_TAB in the SQL database after store number change.
# ===================================================================================================

function Update_SQL_Tables_For_Store_Number_Change
{
	param (
		[string]$storeNumber # New store number to apply
	)
	
	# =============================
	# TABLE NAMES
	# =============================
	$stdTableName = "STD_TAB"
	$terTableName = "TER_TAB"
	
	# =============================
	# SQL FOR STD_TAB
	# =============================
	
	# Create view (required by Storeman load process)
	$createViewCommandStd = @"
IF OBJECT_ID('Std_Load', 'V') IS NOT NULL DROP VIEW Std_Load;
EXEC(N'
CREATE VIEW Std_Load AS
SELECT F1056
FROM $stdTableName;
');
"@
	
	# Update STD_TAB.F1056 to the new store number
	$updateStdTabCommand = @"
UPDATE $stdTableName
SET F1056 = '$storeNumber';
"@
	
	# Drop the view
	$dropViewCommandStd = "DROP VIEW Std_Load;"
	
	
	# =============================
	# SQL FOR TER_TAB
	# =============================
	# Only update rows where F1057 = '901'
	# This ensures ONLY the back-office record gets changed
	# and does NOT break lane entries.
	# =============================
	$updateTerTabCommand = @"
UPDATE $terTableName
SET F1056 = '$storeNumber'
WHERE F1057 = '901';
"@
	
	
	# =============================
	# EXECUTION PIPELINE
	# =============================
	$sqlCommands = @(
		$createViewCommandStd,
		$updateStdTabCommand,
		$dropViewCommandStd,
		$updateTerTabCommand # <-- NEW COMMAND
	)
	
	$allSqlSuccessful = $true
	$failedSqlCommands = @()
	
	foreach ($command in $sqlCommands)
	{
		# Execute_SQL_Commands is part of your existing framework
		# so no changes here.
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$allSqlSuccessful = $false
			$failureMessage = if (-not [string]::IsNullOrWhiteSpace($script:LastSqlExecutionError)) { $script:LastSqlExecutionError } else { "Unknown SQL execution failure." }
			$failedSqlCommands += ("Error: " + $failureMessage + "`r`nSQL:`r`n" + $command)
		}
	}
	
	# =============================
	# RETURN STRUCTURED RESULT
	# =============================
	return @{
		Success	       = $allSqlSuccessful
		FailedCommands = $failedSqlCommands
	}
}

# ===================================================================================================
#                           FUNCTION: Resolve_Startup_Ini_Path
# ---------------------------------------------------------------------------------------------------
# Description:
# Resolves the active Startup.ini path using an explicit value first, then existing script/global
# variables, and finally BasePath when available.
# ===================================================================================================

function Resolve_Startup_Ini_Path
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	$resolvedPath = $StartupIniPath
	
	try
	{
		if ([string]::IsNullOrWhiteSpace($resolvedPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name StartupIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value))
				{
					$resolvedPath = [string]$v.Value
					break
				}
			}
		}
		
		if ([string]::IsNullOrWhiteSpace($resolvedPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name BasePath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value))
				{
					$candidatePath = Join-Path ([string]$v.Value) "Startup.ini"
					if (Test-Path $candidatePath)
					{
						$resolvedPath = $candidatePath
						break
					}
				}
			}
		}
		
		if (-not [string]::IsNullOrWhiteSpace($resolvedPath) -and (Test-Path $resolvedPath))
		{
			return $resolvedPath
		}
	}
	catch { }
	
	return $null
}

# ===================================================================================================
#                           FUNCTION: Get_Machine_Name_Context
# ---------------------------------------------------------------------------------------------------
# Description:
# Normalizes a requested machine name and derives the terminal behavior used by INI + SQL sync paths.
# ===================================================================================================

function Get_Machine_Name_Context
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$CurrentMachineName = $env:COMPUTERNAME,
		[Parameter(Mandatory = $false)]
		[ValidatePattern('^\d{3}$')]
		[string]$ServerTerminal = '901'
	)
	
	$normalizedHost = [string]$MachineName
	if ($null -eq $normalizedHost) { $normalizedHost = "" }
	$normalizedHost = $normalizedHost.Trim()
	$normalizedHost = $normalizedHost -replace '^[\\\/]+', ''
	if ($normalizedHost -match '[\\\/]') { $normalizedHost = ($normalizedHost -split '[\\\/]')[0] }
	if ($normalizedHost -match '\.') { $normalizedHost = ($normalizedHost -split '\.')[0] }
	$normalizedHost = $normalizedHost.Trim().ToUpper()
	
	$currentHost = [string]$CurrentMachineName
	if ($null -eq $currentHost) { $currentHost = "" }
	$currentHost = $currentHost.Trim()
	$currentHost = $currentHost -replace '^[\\\/]+', ''
	if ($currentHost -match '[\\\/]') { $currentHost = ($currentHost -split '[\\\/]')[0] }
	if ($currentHost -match '\.') { $currentHost = ($currentHost -split '\.')[0] }
	$currentHost = $currentHost.Trim().ToUpper()
	
	$nodeRole = "Unknown"
	try
	{
		if (-not [string]::IsNullOrWhiteSpace([string]$script:NodeRole))
		{
			$nodeRole = ([string]$script:NodeRole).Trim()
		}
	}
	catch { }
	
	$isServerNodeRole = ($nodeRole -match '^(?i)(StoreServer|HostServer)$')
	$isLaneNodeRole = ($nodeRole -match '^(?i)Lane$')
	$machineNameFormatText = if ($isServerNodeRole)
	{
		"Valid examples:`n  SERVER001`n  SERVER-001`n  0231SERVER001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	elseif ($isLaneNodeRole)
	{
		"Valid examples:`n  POS001`n  POS-001`n  LN026001`n  0384LANE001`n  001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	else
	{
		"Lane examples:`n  POS001`n  POS-001`n  LN026001`n  0384LANE001`n  001`n`nServer examples:`n  SERVER001`n  SERVER-001`n  0231SERVER001`n`nRule: [letters] + [optional store digits before or after the prefix] + [optional hyphen] + [3-digit terminal]."
	}
	$namePrefix = $null
	$nameSeparator = ""
	$storePrefixRaw = $null
	$storePlacement = "None"
	$terminalRaw = $null
	$inputWasTerminalOnly = $false
	
	if ([string]::IsNullOrWhiteSpace($normalizedHost))
	{
		return [PSCustomObject]@{
			Success            = $false
			ErrorMessage       = "Machine name cannot be empty."
			MachineName        = $null
			CurrentMachineName = $currentHost
			IsSameMachineName  = $false
			IsServerMoniker    = $false
			MachineNumber      = $null
			NodeRole           = $nodeRole
			NamePrefix         = $null
			NameSeparator      = ""
			StorePrefix        = $null
			StorePlacement     = "None"
			TerminalNumber     = $null
			InputWasTerminalOnly = $false
		}
	}
	
	if ($normalizedHost -match '^\d{3}$')
	{
		if ($isServerNodeRole)
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input is only supported for lane-style machine names."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $true
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $null
				NameSeparator        = ""
				StorePrefix          = $null
				StorePlacement       = "None"
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		$currentStore = $null
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('StoreNumber') -and $script:FunctionResults['StoreNumber'])
		{
			$currentStore = ($script:FunctionResults['StoreNumber'].ToString()).Trim()
		}
		else
		{
			$currentStore = Get_Store_Number_From_INI
			if ($currentStore) { $currentStore = ($currentStore.ToString()).Trim() }
		}
		
		if ($currentHost -match '^(?<StorePrefix>\d{3,4})(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$storePrefixRaw = $matches['StorePrefix']
			$namePrefix = $matches['NamePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "Leading"
		}
		elseif ($currentHost -match '^(?<NamePrefix>[A-Z]+)(?<StorePrefix>\d{3,4})(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$namePrefix = $matches['NamePrefix']
			$storePrefixRaw = $matches['StorePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "Embedded"
		}
		elseif ($currentHost -match '^(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
		{
			$namePrefix = $matches['NamePrefix']
			$nameSeparator = $matches['NameSeparator']
			$storePlacement = "None"
		}
		else
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input requires the current machine name to already use a lane naming style with a reusable prefix."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $null
				NameSeparator        = ""
				StorePrefix          = $null
				StorePlacement       = "None"
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		if (($storePlacement -ne "None") -and [string]::IsNullOrWhiteSpace($storePrefixRaw) -and ($currentStore -match '^\d{3,4}$'))
		{
			$storePrefixRaw = $currentStore
		}
		
		if (($storePlacement -ne "None") -and [string]::IsNullOrWhiteSpace($storePrefixRaw))
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Terminal-only input could not determine which store number to preserve from the current naming style."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $namePrefix
				NameSeparator        = $nameSeparator
				StorePrefix          = $null
				StorePlacement       = $storePlacement
				TerminalNumber       = $normalizedHost
				InputWasTerminalOnly = $true
			}
		}
		
		$terminalRaw = $normalizedHost
		$inputWasTerminalOnly = $true
		
		switch ($storePlacement)
		{
			"Leading"  { $normalizedHost = "$storePrefixRaw$namePrefix$nameSeparator$terminalRaw" }
			"Embedded" { $normalizedHost = "$namePrefix$storePrefixRaw$nameSeparator$terminalRaw" }
			default    { $normalizedHost = "$namePrefix$nameSeparator$terminalRaw" }
		}
	}
	elseif ($normalizedHost -match '^(?<StorePrefix>\d{3,4})(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$storePrefixRaw = $matches['StorePrefix']
		$namePrefix = $matches['NamePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "Leading"
	}
	elseif ($normalizedHost -match '^(?<NamePrefix>[A-Z]+)(?<StorePrefix>\d{3,4})(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$namePrefix = $matches['NamePrefix']
		$storePrefixRaw = $matches['StorePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "Embedded"
	}
	elseif ($normalizedHost -match '^(?<NamePrefix>[A-Z]+)(?<NameSeparator>-?)(?<Terminal>\d{3})$')
	{
		$namePrefix = $matches['NamePrefix']
		$nameSeparator = $matches['NameSeparator']
		$terminalRaw = $matches['Terminal']
		$storePlacement = "None"
	}
	else
	{
		return [PSCustomObject]@{
			Success              = $false
			ErrorMessage         = "Invalid format.`n`n$machineNameFormatText"
			MachineName          = $normalizedHost
			CurrentMachineName   = $currentHost
			IsSameMachineName    = $false
			IsServerMoniker      = $isServerNodeRole
			MachineNumber        = $null
			NodeRole             = $nodeRole
			NamePrefix           = $null
			NameSeparator        = ""
			StorePrefix          = $null
			StorePlacement       = "None"
			TerminalNumber       = $null
			InputWasTerminalOnly = $false
		}
	}
	
	if ($normalizedHost.Length -gt 15)
	{
		return [PSCustomObject]@{
			Success              = $false
			ErrorMessage         = "Machine name '$normalizedHost' exceeds the Windows 15-character computer-name limit."
			MachineName          = $normalizedHost
			CurrentMachineName   = $currentHost
			IsSameMachineName    = $false
			IsServerMoniker      = $isServerNodeRole
			MachineNumber        = $null
			NodeRole             = $nodeRole
			NamePrefix           = $namePrefix
			NameSeparator        = $nameSeparator
			StorePrefix          = $storePrefixRaw
			StorePlacement       = $storePlacement
			TerminalNumber       = $terminalRaw
			InputWasTerminalOnly = $inputWasTerminalOnly
		}
	}
	
	$isServerMoniker = $isServerNodeRole
	if ($normalizedHost -match '(?i)SERVER')
	{
		$isServerMoniker = $true
	}
	
	$machineNumber = $null
	if ($isServerMoniker)
	{
		$machineNumber = $ServerTerminal
	}
	else
	{
		$machineNumber = ([int]$terminalRaw).ToString("D3")
		if ($machineNumber -eq '000')
		{
			return [PSCustomObject]@{
				Success              = $false
				ErrorMessage         = "Invalid terminal number extracted from '$normalizedHost' (000 is not allowed)."
				MachineName          = $normalizedHost
				CurrentMachineName   = $currentHost
				IsSameMachineName    = $false
				IsServerMoniker      = $false
				MachineNumber        = $null
				NodeRole             = $nodeRole
				NamePrefix           = $namePrefix
				NameSeparator        = $nameSeparator
				StorePrefix          = $storePrefixRaw
				StorePlacement       = $storePlacement
				TerminalNumber       = $terminalRaw
				InputWasTerminalOnly = $inputWasTerminalOnly
			}
		}
	}
	
	return [PSCustomObject]@{
		Success              = $true
		ErrorMessage         = $null
		MachineName          = $normalizedHost
		CurrentMachineName   = $currentHost
		IsSameMachineName    = ($normalizedHost -eq $currentHost)
		IsServerMoniker      = $isServerMoniker
		MachineNumber        = $machineNumber
		NodeRole             = $nodeRole
		NamePrefix           = $namePrefix
		NameSeparator        = $nameSeparator
		StorePrefix          = $storePrefixRaw
		StorePlacement       = $storePlacement
		TerminalNumber       = $terminalRaw
		InputWasTerminalOnly = $inputWasTerminalOnly
	}
}

# ===================================================================================================
#                          FUNCTION: Update_Startup_Ini_For_Machine
# ---------------------------------------------------------------------------------------------------
# Description:
# Applies machine-specific Startup.ini changes that should happen both after a rename and during a
# same-name sync. This keeps the DBSERVER host update in one place and preserves encoding/newlines.
# ===================================================================================================

function Update_Startup_Ini_For_Machine
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[psobject]$MachineContext,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath
	)
	
	$resolvedStartupIniPath = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
	if ([string]::IsNullOrWhiteSpace($resolvedStartupIniPath))
	{
		return [PSCustomObject]@{
			Success       = $false
			Path          = $null
			DbServerValue = $null
			TerValue      = $null
			Details       = "Startup.ini path could not be resolved."
			ErrorMessage  = "startup.ini not found."
		}
	}
	
	try
	{
		$bytes = [System.IO.File]::ReadAllBytes($resolvedStartupIniPath)
		
		$encStartup = $null
		if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
		{ $encStartup = New-Object System.Text.UTF8Encoding($true) }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
		{ $encStartup = [System.Text.Encoding]::Unicode }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
		{ $encStartup = [System.Text.Encoding]::BigEndianUnicode }
		else
		{ $encStartup = [System.Text.Encoding]::Default }
		
		$textStartup = $encStartup.GetString($bytes)
		$nlStartup = "`r`n"
		if ($textStartup -notmatch "`r`n" -and $textStartup -match "`n") { $nlStartup = "`n" }
		
		$startupLines = $textStartup -split "\r?\n", -1
		
		$existingDbLine = $null
		foreach ($line in $startupLines)
		{
			if ($line -match '^\s*(?i:DBSERVER)\s*=')
			{
				$existingDbLine = $line
				break
			}
		}
		
		$existingDbServer = $null
		if ($existingDbLine -match '^\s*(?i:DBSERVER)\s*=\s*(.+?)\s*$')
		{
			$existingDbServer = $matches[1].Trim()
		}
		
		$instance = $null
		if ($existingDbServer -and ($existingDbServer -match '\\'))
		{
			$parts = $existingDbServer.Split('\')
			if ($parts.Count -ge 2 -and $parts[1]) { $instance = $parts[1] }
		}
		
		$dbServerValue = if ($instance)
		{
			"DBSERVER=$($MachineContext.MachineName)\$instance"
		}
		else
		{
			"DBSERVER=$($MachineContext.MachineName)"
		}
		
		$terValue = $null
		if (-not $MachineContext.IsServerMoniker -and $MachineContext.MachineNumber)
		{
			$terValue = "TER=$($MachineContext.MachineNumber)"
		}
		
		$updatedLines = @()
		$dbServerUpdated = $false
		$terUpdated = $false
		
		foreach ($line in $startupLines)
		{
			$updatedLine = $line
			
			if (-not $MachineContext.IsServerMoniker -and $terValue -and ($updatedLine -match '^\s*(?i:TER)\s*='))
			{
				$updatedLine = $terValue
				$terUpdated = $true
			}
			
			if ($updatedLine -match '^\s*(?i:DBSERVER)\s*=')
			{
				$updatedLine = $dbServerValue
				$dbServerUpdated = $true
			}
			
			$updatedLines += $updatedLine
		}
		
		if (-not $dbServerUpdated)
		{
			$updatedLines += $dbServerValue
			$dbServerUpdated = $true
		}
		
		[System.IO.File]::WriteAllText($resolvedStartupIniPath, ($updatedLines -join $nlStartup), $encStartup)
		
		$detailParts = @()
		$detailParts += "DBSERVER set to '$dbServerValue'."
		if ($MachineContext.IsServerMoniker)
		{
			$detailParts += "TER left unchanged due to SERVER moniker."
		}
		elseif ($terUpdated)
		{
			$detailParts += "TER set to '$terValue'."
		}
		else
		{
			$detailParts += "TER line was not present, so only DBSERVER was enforced here."
		}
		
		return [PSCustomObject]@{
			Success       = $true
			Path          = $resolvedStartupIniPath
			DbServerValue = $dbServerValue
			TerValue      = $terValue
			Details       = ($detailParts -join ' ')
			ErrorMessage  = $null
		}
	}
	catch
	{
		return [PSCustomObject]@{
			Success       = $false
			Path          = $resolvedStartupIniPath
			DbServerValue = $null
			TerValue      = $null
			Details       = "Startup.ini machine sync failed."
			ErrorMessage  = $_.Exception.Message
		}
	}
}

# ===================================================================================================
#                            FUNCTION: Invoke_Machine_Name_Sync
# ---------------------------------------------------------------------------------------------------
# Description:
# Runs the best-effort INI + Startup.ini + SQL sync for a machine name target. This is used after an
# actual rename and also when the requested name already matches the current machine name.
# ===================================================================================================

function Invoke_Machine_Name_Sync
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName = $env:COMPUTERNAME,
		[Parameter(Mandatory = $false)]
		[psobject]$MachineContext,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath,
		[Parameter(Mandatory = $false)]
		[hashtable]$OperationStatus
	)
	
	if (-not $MachineContext -or -not $MachineContext.Success)
	{
		$MachineContext = Get_Machine_Name_Context -MachineName $MachineName -CurrentMachineName $OldMachineName
	}
	
	if (-not $MachineContext.Success)
	{
		if ($OperationStatus -and $OperationStatus.ContainsKey("StartupIniUpdate"))
		{
			$OperationStatus["StartupIniUpdate"].Status = "Failed"
			$OperationStatus["StartupIniUpdate"].Message = "INI sync could not start."
			$OperationStatus["StartupIniUpdate"].Details = $MachineContext.ErrorMessage
		}
		if ($OperationStatus -and $OperationStatus.ContainsKey("SQLDatabaseUpdate"))
		{
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL sync could not start."
			$OperationStatus["SQLDatabaseUpdate"].Details = $MachineContext.ErrorMessage
		}
		
		return [PSCustomObject]@{
			Success             = $false
			MachineContext      = $MachineContext
			IniSyncSuccess      = $false
			StartupIniSuccess   = $false
			SqlSyncSuccess      = $false
			StartupIniPath      = $null
			IniSyncError        = $MachineContext.ErrorMessage
			StartupIniError     = $null
			SqlSyncError        = $MachineContext.ErrorMessage
			StartupIniDetails   = $null
			SqlResult           = $null
		}
	}
	
	$resolvedStartupIniPath = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
	$null = Ensure_Database_Connection_Context -StartupIniPath $resolvedStartupIniPath
	
	$configuredDbServer = $null
	if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConfiguredDBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['ConfiguredDBSERVER']))
	{
		$configuredDbServer = [string]$script:FunctionResults['ConfiguredDBSERVER']
	}
	elseif ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBSERVER']))
	{
		$configuredDbServer = [string]$script:FunctionResults['DBSERVER']
	}
	
	$runtimeDbServer = $null
	if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('RuntimeDBSERVER') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['RuntimeDBSERVER']))
	{
		$runtimeDbServer = [string]$script:FunctionResults['RuntimeDBSERVER']
	}
	
	$iniSyncSuccess = $false
	$iniSyncError = $null
	try
	{
		$iniSyncSuccess = [bool](Update_INIs -newStoreNumber $StoreNumber -MachineName $MachineContext.MachineName -OldStoreNumber $StoreNumber -OldMachineName $OldMachineName -CreateSmsHttpsIfMissing -StartupIniPath $resolvedStartupIniPath)
		if (-not $iniSyncSuccess)
		{
			$iniSyncError = "Update_INIs reported a failure."
		}
	}
	catch
	{
		$iniSyncError = $_.Exception.Message
	}
	
	$startupIniResult = Update_Startup_Ini_For_Machine -MachineContext $MachineContext -StartupIniPath $resolvedStartupIniPath
	$startupIniSuccess = $startupIniResult.Success
	$startupIniError = $startupIniResult.ErrorMessage
	
	if ($OperationStatus -and $OperationStatus.ContainsKey("StartupIniUpdate"))
	{
		if ($iniSyncSuccess -and $startupIniSuccess)
		{
			$OperationStatus["StartupIniUpdate"].Status = "Successful"
			$OperationStatus["StartupIniUpdate"].Message = "INI files updated successfully."
			$OperationStatus["StartupIniUpdate"].Details = $startupIniResult.Details
		}
		else
		{
			$detailParts = @()
			if ($iniSyncError) { $detailParts += $iniSyncError }
			if ($startupIniError) { $detailParts += $startupIniError }
			if ($startupIniResult.Details) { $detailParts += $startupIniResult.Details }
			
			$OperationStatus["StartupIniUpdate"].Status = "Failed"
			$OperationStatus["StartupIniUpdate"].Message = "INI sync completed with errors."
			$OperationStatus["StartupIniUpdate"].Details = ($detailParts -join " ")
		}
	}
	
	$sqlUpdateResult = $null
	$sqlSyncSuccess = $false
	$sqlSyncError = $null
	try
	{
		$sqlUpdateResult = Update_SQL_Tables_For_Machine_Name_Change `
															 -storeNumber $StoreNumber `
															 -machineName $MachineContext.MachineName `
															 -machineNumber $MachineContext.MachineNumber `
															 -OldMachineName $OldMachineName
		
		if ($sqlUpdateResult -and $sqlUpdateResult.Success)
		{
			$sqlSyncSuccess = $true
		}
		else
		{
			$sqlSyncError = "SQL update result indicated failure."
		}
	}
	catch
	{
		$sqlSyncError = $_.Exception.Message
	}
	
	if ($OperationStatus -and $OperationStatus.ContainsKey("SQLDatabaseUpdate"))
	{
		if ($sqlSyncSuccess)
		{
			$detailParts = @("Protected SQL sync completed.")
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RunTabUpdated'))
			{
				$detailParts += "RUN updated: $($sqlUpdateResult.RunTabUpdated)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('LocalTerminalDB'))
			{
				$detailParts += "Local terminal DB: $($sqlUpdateResult.LocalTerminalDB)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('ConfiguredDBSERVER') -and $sqlUpdateResult.ConfiguredDBSERVER)
			{
				$detailParts += "Configured DBSERVER: $($sqlUpdateResult.ConfiguredDBSERVER)."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RuntimeDBSERVER') -and $sqlUpdateResult.RuntimeDBSERVER)
			{
				$detailParts += "Runtime SQL target: $($sqlUpdateResult.RuntimeDBSERVER)."
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Successful"
			$OperationStatus["SQLDatabaseUpdate"].Message = "SQL tables updated successfully after machine sync."
			$OperationStatus["SQLDatabaseUpdate"].Details = ($detailParts -join " ")
		}
		else
		{
			$detailParts = @()
			if ($sqlSyncError) { $detailParts += $sqlSyncError }
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('ConfiguredDBSERVER') -and $sqlUpdateResult.ConfiguredDBSERVER)
			{
				$detailParts += "Configured DBSERVER: $($sqlUpdateResult.ConfiguredDBSERVER)."
			}
			elseif ($configuredDbServer)
			{
				$detailParts += "Configured DBSERVER: $configuredDbServer."
			}
			if ($sqlUpdateResult -and $sqlUpdateResult.ContainsKey('RuntimeDBSERVER') -and $sqlUpdateResult.RuntimeDBSERVER)
			{
				$detailParts += "Runtime SQL target: $($sqlUpdateResult.RuntimeDBSERVER)."
			}
			elseif ($runtimeDbServer)
			{
				$detailParts += "Runtime SQL target: $runtimeDbServer."
			}
			
			$OperationStatus["SQLDatabaseUpdate"].Status = "Failed"
			$OperationStatus["SQLDatabaseUpdate"].Message = "Failed to update SQL tables after machine sync."
			$OperationStatus["SQLDatabaseUpdate"].Details = ($detailParts -join " ")
		}
	}
	
	return [PSCustomObject]@{
		Success             = ($iniSyncSuccess -and $startupIniSuccess -and $sqlSyncSuccess)
		MachineContext      = $MachineContext
		IniSyncSuccess      = $iniSyncSuccess
		StartupIniSuccess   = $startupIniSuccess
		SqlSyncSuccess      = $sqlSyncSuccess
		StartupIniPath      = $resolvedStartupIniPath
		IniSyncError        = $iniSyncError
		StartupIniError     = $startupIniError
		SqlSyncError        = $sqlSyncError
		StartupIniDetails   = $startupIniResult.Details
		SqlResult           = $sqlUpdateResult
		ConfiguredDbServer  = $configuredDbServer
		RuntimeDbServer     = $runtimeDbServer
	}
}

# ===================================================================================================
# 									FUNCTION: Update_INIs
# ---------------------------------------------------------------------------------------------------
# Description:
# Updates ALL required INI files with the new store number.
# Supports both 3-digit and 4-digit store numbers.
# NOTE: No backups are created.
#
# Always-best-effort behavior:
# - Does NOT prompt
# - If MachineName missing -> uses $env:COMPUTERNAME
# - If terminal digits cannot be extracted -> still updates other INIs; skips TER insertion where unsafe
# - If old store cannot be detected -> still forces STORE= and REDIR lines; skips token-replace mapping
# - Startup.ini / Server.ini / SMSStart.ini / INFO_*901_WIN.ini preserve ORIGINAL encoding + ORIGINAL newline style
# - Updates Server.ini when present (same folder as Startup.ini first, then common roots)
#
# IMPORTANT SAFETY RULE:
# - If MachineName contains "SERVER" (anywhere, case-insensitive), we MUST NOT change terminal numbers.
#   That means:
#     * Do NOT derive terminal from name
#     * Do NOT rewrite TER= lines in Startup.ini / Server.ini / SMSStart.ini
#     * DO still update STORE= and REDIRMAIL/REDIRMSG to <STORE>901
#
# SmsHttps.INI behavior:
# - SERVER moniker:
#     * KEEP multiple processor entries
#     * ONLY update the store prefix in:
#         - processor key (e.g., 0231901 -> 0242901)
#         - REDIRMAIL/REDIRMSG values (e.g., 0231901 -> 0242901)
#     * Clear LicenseGUID (always)
# - NON-SERVER moniker:
#     * (Safe default here) we DO NOT touch processors (to avoid accidental changes)
#     * Clear LicenseGUID (always)
#
# NEW:
# - SMSStart.ini may contain PARAMETERS=/ini=Startup902 (or Startup902.ini).
#   We resolve that referenced INI and update it too (Store/Redir/token-replace).
#   We DO NOT force TER changes in the referenced Startup*.ini (we generally leave TER as-is).
# ===================================================================================================

function Update_INIs
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$newStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$MachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldStoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName,
		[Parameter(Mandatory = $false)]
		[string]$StartupIniPath,
		[Parameter(Mandatory = $false)]
		[string]$GlobalSmsStartIniPath,
		[Parameter(Mandatory = $false)]
		[string]$WinIniPath,
		[Parameter(Mandatory = $false)]
		[string]$SmsHttpsIniPath,
		[Parameter(Mandatory = $false)]
		[switch]$CreateSmsHttpsIfMissing
	)
	
	$success = $true
	
	# ------------------------------------------------------------------------------------------------
	# Prefer already-built variables for paths (no hardcoded probing)
	# ------------------------------------------------------------------------------------------------
	try
	{
		if ([string]::IsNullOrWhiteSpace($StartupIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name StartupIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $StartupIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($GlobalSmsStartIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name GlobalSmsStartIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $GlobalSmsStartIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($WinIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name WinIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $WinIniPath = [string]$v.Value; break }
			}
		}
		if ([string]::IsNullOrWhiteSpace($SmsHttpsIniPath))
		{
			foreach ($sc in @('Script', 'Global'))
			{
				$v = Get-Variable -Name SmsHttpsIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v.Value)) { $SmsHttpsIniPath = [string]$v.Value; break }
			}
		}
	}
	catch { }
	
	# Optional: reuse already-built core paths if present (does not rebuild new ones)
	$BasePath = $null
	$DbsPath = $null
	$SystemIniVar = $null
	try
	{
		foreach ($sc in @('Script', 'Global'))
		{
			if (-not $BasePath)
			{
				$v = Get-Variable -Name BasePath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $BasePath = [string]$v.Value }
			}
			if (-not $DbsPath)
			{
				$v = Get-Variable -Name DbsPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $DbsPath = [string]$v.Value }
			}
			if (-not $SystemIniVar)
			{
				$v = Get-Variable -Name SystemIniPath -Scope $sc -ErrorAction SilentlyContinue
				if ($v -and $v.Value) { $SystemIniVar = [string]$v.Value }
			}
		}
	}
	catch { }
	
	# If caller didn't supply MachineName, use current machine name (no prompt, no rename)
	if ([string]::IsNullOrWhiteSpace($MachineName))
	{
		$MachineName = $env:COMPUTERNAME
	}
	
	# ------------------------------------------------------------------------------------------------
	# Normalize NEW store (preserve user width for STORE=/REDIR lines)
	# ------------------------------------------------------------------------------------------------
	$newStoreTrim = ($newStoreNumber + "").Trim()
	$newStoreInt = [int]$newStoreTrim
	$newStore3 = $newStoreInt.ToString("D3")
	$newStore4 = $newStoreInt.ToString("D4")
	
	# ------------------------------------------------------------------------------------------------
	# Normalize machine name (strip UNC/path + domain) + ALWAYS uppercase
	# ------------------------------------------------------------------------------------------------
	$mn = ($MachineName + "").Trim()
	$mn = $mn -replace '^[\\\/]+', ''
	if ($mn -match '[\\\/]') { $mn = ($mn -split '[\\\/]')[0] }
	if ($mn -match '\.') { $mn = ($mn -split '\.')[0] }
	$mn = $mn.Trim().ToUpper()
	
	# ------------------------------------------------------------------------------------------------
	# SERVER MONIKER RULE: if name contains SERVER anywhere -> NEVER change terminal numbers
	# ------------------------------------------------------------------------------------------------
	$isServerMoniker = $false
	if ($mn -match '(?i)SERVER') { $isServerMoniker = $true }
	elseif ($script:NodeRole -and ($script:NodeRole -match '^(?i)(StoreServer|HostServer)$')) { $isServerMoniker = $true }
	
	# ------------------------------------------------------------------------------------------------
	# Derive terminal ONLY for NON-server moniker machines
	# ------------------------------------------------------------------------------------------------
	$newTerminal = $null
	if (-not $isServerMoniker -and ($mn -match '(\d{1,3})$'))
	{
		$newTerminal = ([int]$Matches[1]).ToString("D3")
		if ($newTerminal -eq "000") { $newTerminal = $null }
	}
	
	# ===============================================================================================
	# 1) Resolve + Update Startup.ini (required)  (no hardcoded probing; must already be set)
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($StartupIniPath) -and $BasePath)
	{
		try
		{
			$p = Join-Path $BasePath "Startup.ini"
			if (Test-Path $p) { $StartupIniPath = $p }
		}
		catch { }
	}
	
	if ([string]::IsNullOrWhiteSpace($StartupIniPath) -or -not (Test-Path $StartupIniPath))
	{
		Write-Host "startup.ini not found. Ensure `$StartupIniPath is already set (or pass -StartupIniPath)." -ForegroundColor Red
		return $false
	}
	
	$oldStoreDetected = $null
	$startupDir = $null
	try { $startupDir = Split-Path -Path $StartupIniPath -Parent }
	catch { $startupDir = $null }
	
	try
	{
		$bytes = [System.IO.File]::ReadAllBytes($StartupIniPath)
		
		$encStartup = $null
		if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
		{ $encStartup = New-Object System.Text.UTF8Encoding($true) }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
		{ $encStartup = [System.Text.Encoding]::Unicode }
		elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
		{ $encStartup = [System.Text.Encoding]::BigEndianUnicode }
		else
		{ $encStartup = [System.Text.Encoding]::Default }
		
		$textStartup = $encStartup.GetString($bytes)
		
		$nlStartup = "`r`n"
		if ($textStartup -notmatch "`r`n" -and $textStartup -match "`n") { $nlStartup = "`n" }
		
		$startupLines = $textStartup -split "\r?\n", -1
		
		# Detect OLD store number (prefer STORE=)
		foreach ($l in $startupLines)
		{
			if ($l -match '^\s*STORE\s*=\s*(\d{3,4})\s*$') { $oldStoreDetected = $matches[1]; break }
		}
		if (-not $oldStoreDetected)
		{
			foreach ($l in $startupLines)
			{
				if ($l -match '^\s*SERVERNAME\s*=\s*(\d{3,4})[A-Za-z]') { $oldStoreDetected = $matches[1]; break }
			}
		}
		if (-not $oldStoreDetected)
		{
			foreach ($l in $startupLines)
			{
				if ($l -match '\\\\(\d{3,4})[A-Za-z]') { $oldStoreDetected = $matches[1]; break }
			}
		}
		
		# If user provided OldStoreNumber, override detection (if valid)
		if (-not [string]::IsNullOrWhiteSpace($OldStoreNumber))
		{
			$os = $OldStoreNumber.Trim()
			if ($os -match '^(?!0+$)\d{3,4}$') { $oldStoreDetected = $os }
		}
		
		$doTokenReplace = $true
		if (-not $oldStoreDetected) { $doTokenReplace = $false }
		
		$oldTokens = @()
		$newTokens = @()
		if ($doTokenReplace)
		{
			$oldStoreInt = [int]$oldStoreDetected
			$oldTokens = @($oldStoreInt.ToString("D4"), $oldStoreInt.ToString("D3"))
			$newTokens = @($newStore4, $newStore3)
		}
		
		for ($i = 0; $i -lt $startupLines.Count; $i++)
		{
			$line = $startupLines[$i]
			
			$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
			$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
			
			# TER only when NON-server and we have a derived terminal
			if (-not $isServerMoniker -and $newTerminal)
			{
				$line = $line -replace '^\s*TER\s*=\s*\d{1,4}\s*$', ("TER=" + $newTerminal)
			}
			
			# Always uppercase SERVERNAME value if present
			if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
			{
				$ws = $matches['ws']
				$val = ($matches['val'] + "").Trim()
				if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
			}
			
			# Token replace mapping (SAFE: do not rewrite terminal suffixes like SERVER001)
			if ($doTokenReplace)
			{
				for ($t = 0; $t -lt $oldTokens.Count; $t++)
				{
					$oldEsc = [regex]::Escape($oldTokens[$t])
					$newTok = $newTokens[$t]
					
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
					$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
					$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
				}
			}
			
			$startupLines[$i] = $line
		}
		
		[System.IO.File]::WriteAllText($StartupIniPath, ($startupLines -join $nlStartup), $encStartup)
		Write-Host "Updated startup.ini" -ForegroundColor Green
		
		if (-not $oldStoreDetected) { $oldStoreDetected = $newStoreTrim }
	}
	catch
	{
		$success = $false
		Write-Host "Failed updating startup.ini: $($_.Exception.Message)" -ForegroundColor Red
		if (-not $oldStoreDetected) { $oldStoreDetected = $newStoreTrim }
	}
	
	# ===============================================================================================
	# 1B) Update Server.ini (optional) (relative to Startup.ini / BasePath only)
	# ===============================================================================================
	$serverIniPath = $null
	try
	{
		if ($startupDir)
		{
			$p = Join-Path $startupDir "Server.ini"
			if (Test-Path $p) { $serverIniPath = $p }
		}
		if (-not $serverIniPath -and $BasePath)
		{
			$p = Join-Path $BasePath "Server.ini"
			if (Test-Path $p) { $serverIniPath = $p }
		}
	}
	catch { $serverIniPath = $null }
	
	if (-not [string]::IsNullOrWhiteSpace($serverIniPath) -and (Test-Path $serverIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($serverIniPath)
			
			$encServer = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encServer = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encServer = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encServer = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encServer = [System.Text.Encoding]::Default }
			
			$textServer = $encServer.GetString($bytes)
			$nlServer = "`r`n"
			if ($textServer -notmatch "`r`n" -and $textServer -match "`n") { $nlServer = "`n" }
			
			$serverLines = $textServer -split "\r?\n", -1
			
			$doTokenReplaceS = $false
			$oldTokensS = @()
			$newTokensS = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplaceS = $true
				$osi = [int]$oldStoreDetected
				$oldTokensS = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokensS = @($newStore4, $newStore3)
			}
			
			for ($i = 0; $i -lt $serverLines.Count; $i++)
			{
				$line = $serverLines[$i]
				
				$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
				$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
				
				if (-not $isServerMoniker -and $newTerminal)
				{
					$line = $line -replace '^\s*TER\s*=\s*\d{1,4}\s*$', ("TER=" + $newTerminal)
				}
				
				if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
				{
					$ws = $matches['ws']
					$val = ($matches['val'] + "").Trim()
					if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
				}
				
				if ($doTokenReplaceS)
				{
					for ($t = 0; $t -lt $oldTokensS.Count; $t++)
					{
						$oldEsc = [regex]::Escape($oldTokensS[$t])
						$newTok = $newTokensS[$t]
						
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
						$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
					}
				}
				
				$serverLines[$i] = $line
			}
			
			[System.IO.File]::WriteAllText($serverIniPath, ($serverLines -join $nlServer), $encServer)
			Write-Host "Updated Server.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating Server.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 1C) Update Office\System.ini (optional) using existing $SystemIniPath if available
	# ===============================================================================================
	$systemIniPath = $null
	if (-not [string]::IsNullOrWhiteSpace($SystemIniVar) -and (Test-Path $SystemIniVar))
	{
		$systemIniPath = $SystemIniVar
	}
	elseif ($startupDir)
	{
		$p = Join-Path $startupDir "Office\System.ini"
		if (Test-Path $p) { $systemIniPath = $p }
	}
	
	if (-not [string]::IsNullOrWhiteSpace($systemIniPath) -and (Test-Path $systemIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($systemIniPath)
			
			$encSys = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encSys = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encSys = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encSys = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encSys = [System.Text.Encoding]::Default }
			
			$textSys = $encSys.GetString($bytes)
			
			$nlSys = "`r`n"
			if ($textSys -notmatch "`r`n" -and $textSys -match "`n") { $nlSys = "`n" }
			
			$linesSys = $textSys -split "\r?\n", -1
			
			$newStoreIsOne = $false
			try { if ([int]$newStoreTrim -eq 1) { $newStoreIsOne = $true } }
			catch { $newStoreIsOne = $false }
			
			$oldTokensSys = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$osi = [int]$oldStoreDetected
				$oldTokensSys = @($osi.ToString("D4"), $osi.ToString("D3"))
			}
			
			$inSmsSection = $false
			$changedSys = $false
			
			for ($i = 0; $i -lt $linesSys.Count; $i++)
			{
				$line = $linesSys[$i]
				
				if ($line -match '^\s*\[(?<sec>.+?)\]\s*$')
				{
					$sec = $matches['sec'].Trim()
					$inSmsSection = ($sec -ieq 'SMS')
					continue
				}
				
				if (-not $inSmsSection) { continue }
				
				if ($line -match '^(?<ws>\s*)(?i:Name)\s*=\s*(?<val>.*)\s*$')
				{
					$ws = $matches['ws']
					$val = $matches['val']
					$origVal = $val
					
					$val2 = ($val + "").Trim()
					
					if ($newStoreIsOne)
					{
						$val2 = [regex]::Replace($val2, '\s+(?:0001|001|01|1)\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
						foreach ($tok in @('0001', '001', '01', '1'))
						{
							$val2 = [regex]::Replace($val2, '\s+' + [regex]::Escape($tok) + '\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
						}
						$val2 = $val2.Trim()
					}
					else
					{
						$replaced = $false
						foreach ($tok in $oldTokensSys)
						{
							if (-not [string]::IsNullOrWhiteSpace($tok) -and ($val2 -match ("(?<!\d)" + [regex]::Escape($tok) + "(?!\d)")))
							{
								$val2 = [regex]::Replace($val2, "(?<!\d)" + [regex]::Escape($tok) + "(?!\d)", $newStoreTrim)
								$replaced = $true
							}
						}
						
						if (-not $replaced -and ($val2 -match '\s+\d{3,4}\s*$'))
						{
							$val2 = [regex]::Replace($val2, '\s+\d{3,4}\s*$', (" " + $newStoreTrim))
							$replaced = $true
						}
						
						if (-not $replaced)
						{
							$val2 = ($val2.TrimEnd() + " " + $newStoreTrim).Trim()
						}
					}
					
					if ($val2 -ne $origVal)
					{
						$linesSys[$i] = ($ws + "Name=" + $val2)
						$changedSys = $true
					}
					
					break
				}
			}
			
			if ($changedSys)
			{
				[System.IO.File]::WriteAllText($systemIniPath, ($linesSys -join $nlSys), $encSys)
				Write-Host "Updated Office\System.ini (SMS Name)" -ForegroundColor Green
			}
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating Office\System.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 2) Update Global SMSStart.ini (optional)
	#     + NEW: detect PARAMETERS=/ini=Startup### and update referenced INI too
	#     + NEW: update [TIMESYNC] SERVER=... (store prefix) + ENABLE=1 (insert if missing)
	# ===============================================================================================
	if (-not [string]::IsNullOrWhiteSpace($GlobalSmsStartIniPath) -and (Test-Path $GlobalSmsStartIniPath))
	{
		$referencedIniNames = @()
		
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($GlobalSmsStartIniPath)
			
			$encSmsStart = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encSmsStart = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encSmsStart = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encSmsStart = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encSmsStart = [System.Text.Encoding]::Default }
			
			$textSmsStart = $encSmsStart.GetString($bytes)
			$nlSmsStart = "`r`n"
			if ($textSmsStart -notmatch "`r`n" -and $textSmsStart -match "`n") { $nlSmsStart = "`n" }
			
			$globalLines = $textSmsStart -split "\r?\n", -1
			
			$doTokenReplace = $false
			$oldTokens = @()
			$newTokens = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplace = $true
				$osi = [int]$oldStoreDetected
				$oldTokens = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokens = @($newStore4, $newStore3)
			}
			
			$out = @()
			$inSmsStartSection = $false
			$terFound = $false
			$terLine = $null
			if (-not $isServerMoniker -and $newTerminal) { $terLine = "TER=$newTerminal" }
			
			$inTimeSyncSection = $false
			$tsEnableFound = $false
			
			foreach ($raw in $globalLines)
			{
				$line = $raw
				
				# Detect PARAMETERS=/ini=Startup902 (or Startup902.ini)
				if ($line -match '^\s*(?i:PARAMETERS)\s*=\s*(?<p>.*)\s*$')
				{
					$p = $matches['p']
					if (-not [string]::IsNullOrWhiteSpace($p))
					{
						if ($p -match '(?i)(?:^|\s)/ini\s*=\s*(?<ini>[^\s"]+)')
						{
							$iniName = $matches['ini'].Trim().Trim('"').Trim("'")
							if (-not [string]::IsNullOrWhiteSpace($iniName))
							{
								$referencedIniNames += $iniName
							}
						}
					}
				}
				
				if ($line -match '^\s*\[(.+?)\]\s*$')
				{
					# Leaving [SMSSTART] -> ensure TER inserted if needed
					if ($inSmsStartSection -and -not $terFound -and $terLine)
					{
						$out += $terLine
						$terFound = $true
					}
					
					# Leaving [TIMESYNC] -> ensure ENABLE=1 if missing
					if ($inTimeSyncSection -and -not $tsEnableFound)
					{
						$out += "ENABLE=1"
						$tsEnableFound = $true
					}
					
					$sectionName = $matches[1].Trim()
					$inSmsStartSection = ($sectionName -ieq 'SMSSTART')
					if ($inSmsStartSection) { $terFound = $false }
					
					$inTimeSyncSection = ($sectionName -ieq 'TIMESYNC')
					if ($inTimeSyncSection) { $tsEnableFound = $false }
					
					$out += $line
					continue
				}
				
				if ($inSmsStartSection)
				{
					$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
					$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
					
					# TER only when allowed (non-server)
					if ($terLine -and ($line -match '^\s*(?i:TER)\s*='))
					{
						$line = $terLine
						$terFound = $true
					}
					
					if ($doTokenReplace)
					{
						for ($t = 0; $t -lt $oldTokens.Count; $t++)
						{
							$oldEsc = [regex]::Escape($oldTokens[$t])
							$newTok = $newTokens[$t]
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
							$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
							$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
						}
					}
				}
				elseif ($inTimeSyncSection)
				{
					# Force ENABLE=1 (and mark found)
					if ($line -match '^\s*(?i:ENABLE)\s*=')
					{
						$line = "ENABLE=1"
						$tsEnableFound = $true
					}
					
					# Update SERVER= store prefix (and force uppercase)
					if ($line -match '^\s*(?i:SERVER)\s*=\s*(?<sv>.*)\s*$')
					{
						$sv = ($matches['sv'] + "").Trim()
						
						if ($sv -match '^(?<st>\d{3,4})(?<rest>.*)$')
						{
							$st = $matches['st']
							$rest = $matches['rest']
							$newSt = ([int]$newStoreTrim).ToString(("D{0}" -f $st.Length))
							$sv = $newSt + $rest
						}
						else
						{
							if ($doTokenReplace)
							{
								for ($t = 0; $t -lt $oldTokens.Count; $t++)
								{
									$oldEsc = [regex]::Escape($oldTokens[$t])
									$newTok = $newTokens[$t]
									$sv = $sv -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
									$sv = $sv -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
									$sv = $sv -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
								}
							}
						}
						
						$line = "SERVER=" + $sv.ToUpper()
					}
				}
				
				$out += $line
			}
			
			# If file ended while still in sections
			if ($inSmsStartSection -and -not $terFound -and $terLine)
			{
				$out += $terLine
			}
			if ($inTimeSyncSection -and -not $tsEnableFound)
			{
				$out += "ENABLE=1"
			}
			
			[System.IO.File]::WriteAllText($GlobalSmsStartIniPath, ($out -join $nlSmsStart), $encSmsStart)
			Write-Host "Updated SMSStart.ini (SMSSTART + TIMESYNC)" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SMSStart.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
		
		# ---- Update referenced Startup*.ini files (PARAMETERS=/ini=...) ----
		if ($referencedIniNames -and $referencedIniNames.Count -gt 0)
		{
			$smsStartDir = $null
			try { $smsStartDir = Split-Path -Path $GlobalSmsStartIniPath -Parent }
			catch { $smsStartDir = $null }
			
			foreach ($iniRef in ($referencedIniNames | Select-Object -Unique))
			{
				try
				{
					$iniTarget = ($iniRef + "").Trim()
					if ([string]::IsNullOrWhiteSpace($iniTarget)) { continue }
					if ($iniTarget -notmatch '\.ini$') { $iniTarget = ($iniTarget + ".ini") }
					
					$candidatePaths = @()
					
					if ($iniTarget -match '^[A-Za-z]:\\' -or $iniTarget -match '^\\\\')
					{
						$candidatePaths += $iniTarget
					}
					else
					{
						if ($smsStartDir) { $candidatePaths += (Join-Path $smsStartDir $iniTarget) }
						if ($BasePath) { $candidatePaths += (Join-Path $BasePath    $iniTarget) }
						if ($startupDir) { $candidatePaths += (Join-Path $startupDir  $iniTarget) }
					}
					
					$resolved = $null
					foreach ($cp in $candidatePaths)
					{
						if (-not [string]::IsNullOrWhiteSpace($cp) -and (Test-Path $cp))
						{
							$resolved = $cp
							break
						}
					}
					
					if (-not $resolved) { continue }
					
					$b = [System.IO.File]::ReadAllBytes($resolved)
					
					$encRef = $null
					if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
					{ $encRef = New-Object System.Text.UTF8Encoding($true) }
					elseif ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE)
					{ $encRef = [System.Text.Encoding]::Unicode }
					elseif ($b.Length -ge 2 -and $b[0] -eq 0xFE -and $b[1] -eq 0xFF)
					{ $encRef = [System.Text.Encoding]::BigEndianUnicode }
					else
					{ $encRef = [System.Text.Encoding]::Default }
					
					$txt = $encRef.GetString($b)
					$nlRef = "`r`n"
					if ($txt -notmatch "`r`n" -and $txt -match "`n") { $nlRef = "`n" }
					
					$lines = $txt -split "\r?\n", -1
					
					$doTR = $false
					$ot = @()
					$nt = @()
					if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
					{
						$doTR = $true
						$osi = [int]$oldStoreDetected
						$ot = @($osi.ToString("D4"), $osi.ToString("D3"))
						$nt = @($newStore4, $newStore3)
					}
					
					for ($i = 0; $i -lt $lines.Count; $i++)
					{
						$line = $lines[$i]
						
						$line = [regex]::Replace($line, '^\s*(?i:STORE)\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim))
						$line = [regex]::Replace($line, '^\s*(?i:Store)\s*=\s*\d{1,4}\s*$', ("Store=" + $newStoreTrim))
						
						$line = [regex]::Replace(
							$line,
							'^\s*(?i:REDIRMAIL|REDIRMSG|RedirMail|RedirMsg)\s*=\s*\d{1,4}(901)\s*$',
							{
								param ($m)
								$key = $m.Value.Split('=')[0].Trim()
								return ($key + "=" + $newStoreTrim + "901")
							}
						)
						
						if ($line -match '^(?<ws>\s*)(?i:SERVERNAME)\s*=\s*(?<val>.*)\s*$')
						{
							$ws = $matches['ws']
							$val = ($matches['val'] + "").Trim()
							if ($val.Length -gt 0) { $line = $ws + "SERVERNAME=" + $val.ToUpper() }
						}
						
						# Do NOT force TER changes in referenced Startup*.ini
						
						if ($doTR)
						{
							for ($t = 0; $t -lt $ot.Count; $t++)
							{
								$oldEsc = [regex]::Escape($ot[$t])
								$newTok = $nt[$t]
								$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
								$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
								$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
							}
						}
						
						$lines[$i] = $line
					}
					
					[System.IO.File]::WriteAllText($resolved, ($lines -join $nlRef), $encRef)
					Write-Host ("Updated referenced INI: " + $resolved) -ForegroundColor Green
				}
				catch
				{
					$success = $false
					Write-Host ("Failed updating referenced INI '" + $iniRef + "': " + $_.Exception.Message) -ForegroundColor Red
				}
			}
		}
	}
	
	# ===============================================================================================
	# 3) Update INFO_*901_WIN.ini (optional) (use variable; small fallback via $DbsPath only)
	# ===============================================================================================
	if ([string]::IsNullOrWhiteSpace($WinIniPath) -and $DbsPath -and (Test-Path $DbsPath))
	{
		try
		{
			$f = Get-ChildItem -Path $DbsPath -File -Filter "INFO_*901_WIN.ini" -ErrorAction SilentlyContinue | Select-Object -First 1
			if ($f) { $WinIniPath = $f.FullName }
		}
		catch { }
	}
	
	if (-not [string]::IsNullOrWhiteSpace($WinIniPath) -and (Test-Path $WinIniPath))
	{
		try
		{
			$bytes = [System.IO.File]::ReadAllBytes($WinIniPath)
			
			$encWin = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
			{ $encWin = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE)
			{ $encWin = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)
			{ $encWin = [System.Text.Encoding]::BigEndianUnicode }
			else
			{ $encWin = [System.Text.Encoding]::Default }
			
			$textWin = $encWin.GetString($bytes)
			$nlWin = "`r`n"
			if ($textWin -notmatch "`r`n" -and $textWin -match "`n") { $nlWin = "`n" }
			
			$winLines = $textWin -split "\r?\n", -1
			
			$doTokenReplaceW = $false
			$oldTokensW = @()
			$newTokensW = @()
			if (-not [string]::IsNullOrWhiteSpace($oldStoreDetected) -and ($oldStoreDetected -match '^\d{3,4}$'))
			{
				$doTokenReplaceW = $true
				$osi = [int]$oldStoreDetected
				$oldTokensW = @($osi.ToString("D4"), $osi.ToString("D3"))
				$newTokensW = @($newStore4, $newStore3)
			}
			
			for ($j = 0; $j -lt $winLines.Count; $j++)
			{
				$line = $winLines[$j]
				$line = $line -replace '^\s*STORE\s*=\s*\d{1,4}\s*$', ("STORE=" + $newStoreTrim)
				$line = $line -replace '^\s*(REDIRMAIL|REDIRMSG)\s*=\s*\d{1,4}(901)\s*$', ('$1=' + $newStoreTrim + '$2')
				
				if ($doTokenReplaceW)
				{
					for ($t = 0; $t -lt $oldTokensW.Count; $t++)
					{
						$oldEsc = [regex]::Escape($oldTokensW[$t])
						$newTok = $newTokensW[$t]
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=[A-Za-z])"), $newTok
						$line = $line -replace ("(?<!\d)" + $oldEsc + "(?=\d{3})"), $newTok
						$line = $line -replace ("(?<![\dA-Za-z])" + $oldEsc + "(?!\d)"), $newTok
					}
				}
				
				$winLines[$j] = $line
			}
			
			[System.IO.File]::WriteAllText($WinIniPath, ($winLines -join $nlWin), $encWin)
			Write-Host "Updated INFO_*901_WIN.ini" -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating INFO_*901_WIN.ini: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	
	# ===============================================================================================
	# 4) Update SmsHttps.INI (FORCED)
	#   - SERVER moniker:
	#       * KEEP multiple processors
	#       * ONLY update store prefix in processor keys + REDIRMAIL/REDIRMSG values
	#       * Always clear LicenseGUID
	#   - NON-SERVER:
	#       * Enforce EXACTLY ONE processor entry keyed <STORE><TER> (store width inferred from file when possible)
	#       * Purge other processor lines
	#       * Force REDIRMAIL/REDIRMSG=<STORE>901
	#       * Always clear LicenseGUID
	#   - ALWAYS writes the file (even if identical), so it "triggers" every run.
	# ===============================================================================================
	
	# --- Resolve SmsHttpsIniPath if missing (fast, no broad probing) ---
	if ([string]::IsNullOrWhiteSpace($SmsHttpsIniPath))
	{
		try
		{
			$startupDirLocal = $null
			if (-not [string]::IsNullOrWhiteSpace($StartupIniPath))
			{
				try { $startupDirLocal = Split-Path -Path $StartupIniPath -Parent }
				catch { $startupDirLocal = $null }
			}
		}
		catch { }
	}
	if (-not [string]::IsNullOrWhiteSpace($SmsHttpsIniPath) -and (Test-Path $SmsHttpsIniPath))
	{
		try
		{
			# If terminal couldn't be derived from machine name, try TER= in Startup.ini (NON-server only)
			if (-not $isServerMoniker -and -not $newTerminal -and -not [string]::IsNullOrWhiteSpace($StartupIniPath) -and (Test-Path $StartupIniPath))
			{
				try
				{
					$terLine = (Get-Content -Path $StartupIniPath -ErrorAction Stop | Where-Object { $_ -match '^\s*TER\s*=\s*\d{1,4}\s*$' } | Select-Object -First 1)
					if ($terLine -and ($terLine -match '^\s*TER\s*=\s*(\d{1,4})\s*$'))
					{
						$tmpTer = $matches[1]
						if ($tmpTer -match '^\d{1,3}$')
						{
							$newTerminal = ([int]$tmpTer).ToString("D3")
							if ($newTerminal -eq "000") { $newTerminal = $null }
						}
					}
				}
				catch { }
			}
			
			$bytes = [System.IO.File]::ReadAllBytes($SmsHttpsIniPath)
			
			$enc = $null
			if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $enc = New-Object System.Text.UTF8Encoding($true) }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [System.Text.Encoding]::Unicode }
			elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $enc = [System.Text.Encoding]::BigEndianUnicode }
			else { $enc = [System.Text.Encoding]::Default }
			
			$text = $enc.GetString($bytes)
			
			$nl = "`r`n"
			if ($text -notmatch "`r`n" -and $text -match "`n") { $nl = "`n" }
			
			$lines = $text -split "\r?\n", -1
			
			# --- Ensure/clear [GENERAL] LicenseGUID ---
			$genStart = -1
			for ($i = 0; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[\s*GENERAL\s*\]\s*$') { $genStart = $i; break }
			}
			if ($genStart -lt 0)
			{
				$lines = @("[GENERAL]", "LicenseGUID=", "") + $lines
			}
			else
			{
				$genEnd = $lines.Length
				for ($i = $genStart + 1; $i -lt $lines.Length; $i++)
				{
					if ($lines[$i] -match '^\s*\[.*\]\s*$') { $genEnd = $i; break }
				}
				
				$lgFound = $false
				for ($i = $genStart + 1; $i -lt $genEnd; $i++)
				{
					if ($lines[$i] -match '^(?<ws>\s*)(?i:LicenseGUID)\s*=\s*(?<val>.*)\s*$')
					{
						$lines[$i] = ($matches['ws'] + "LicenseGUID=")
						$lgFound = $true
						break
					}
				}
				if (-not $lgFound)
				{
					$pre = @()
					$post = @()
					for ($k = 0; $k -lt $genEnd; $k++) { $pre += $lines[$k] }
					for ($k = $genEnd; $k -lt $lines.Length; $k++) { $post += $lines[$k] }
					$lines = @($pre + "LicenseGUID=" + $post)
				}
			}
			
			# --- Find [PROCESSORS] section (create if missing) ---
			$secStart = -1
			for ($i = 0; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[\s*PROCESSORS?\s*\]\s*$') { $secStart = $i; break }
			}
			
			if ($secStart -lt 0)
			{
				# Always create it (forced behavior)
				$lines = @($lines + "" + "[PROCESSORS]")
				$secStart = ($lines.Length - 1)
			}
			
			$secEnd = $lines.Length
			for ($i = $secStart + 1; $i -lt $lines.Length; $i++)
			{
				if ($lines[$i] -match '^\s*\[.*\]\s*$') { $secEnd = $i; break }
			}
			
			# --- Infer store width from existing processor keys when possible (6=3+3, 7=4+3) ---
			$storeWidth = $newStoreTrim.Length
			try
			{
				for ($i = $secStart + 1; $i -lt $secEnd; $i++)
				{
					if ($lines[$i] -match '^\s*(\d{6,7})\s*=')
					{
						$keyLen = $matches[1].Length
						if ($keyLen -eq 7) { $storeWidth = 4; break }
						if ($keyLen -eq 6) { $storeWidth = 3; break }
					}
				}
			}
			catch { }
			
			$storeNorm = ([int]$newStoreTrim).ToString(("D{0}" -f $storeWidth))
			$redirKey901 = $storeNorm + "901"
			
			# =========================
			# SERVER: keep many processors; update store prefix in key + REDIR store prefix
			# =========================
			if ($isServerMoniker)
			{
				for ($i = $secStart + 1; $i -lt $secEnd; $i++)
				{
					$line = $lines[$i]
					if ($line -match '^(?<ws>\s*)(?<store>\d{3,4})(?<term>\d{3})\s*=\s*(?<rhs>.*)$')
					{
						$ws = $matches['ws']
						$st = $matches['store']
						$term = $matches['term']
						$rhs = $matches['rhs']
						
						$newStorePart = ([int]$newStoreTrim).ToString(("D{0}" -f $st.Length))
						$newKey = $newStorePart + $term
						
						$rhs2 = $rhs
						$rhs2 = [regex]::Replace($rhs2, '(?i)\b(REDIRMAIL)\s*=\s*(\d{3,4})(\d{3})\b', {
								param ($m)
								$k = $m.Groups[1].Value
								$st2 = $m.Groups[2].Value
								$tr2 = $m.Groups[3].Value
								$nst2 = ([int]$newStoreTrim).ToString(("D{0}" -f $st2.Length))
								return ($k + "=" + $nst2 + $tr2)
							})
						$rhs2 = [regex]::Replace($rhs2, '(?i)\b(REDIRMSG)\s*=\s*(\d{3,4})(\d{3})\b', {
								param ($m)
								$k = $m.Groups[1].Value
								$st2 = $m.Groups[2].Value
								$tr2 = $m.Groups[3].Value
								$nst2 = ([int]$newStoreTrim).ToString(("D{0}" -f $st2.Length))
								return ($k + "=" + $nst2 + $tr2)
							})
						
						$lines[$i] = ($ws + $newKey + "=" + $rhs2)
					}
				}
			}
			# =========================
			# NON-SERVER: enforce ONE processor keyed <STORE><TER> and force REDIRs to <STORE>901
			# =========================
			else
			{
				if ($newTerminal)
				{
					$newKey = $storeNorm + $newTerminal
					
					# Collect existing processor entries (numeric)
					$entries = @()
					for ($i = $secStart + 1; $i -lt $secEnd; $i++)
					{
						if ($lines[$i] -match '^(?<ws>\s*)(?<key>\d{6,7})\s*=\s*(?<rhs>.*)$')
						{
							$entries += [pscustomobject]@{
								Index = $i
								Ws    = $matches['ws']
								Key   = $matches['key']
								Rhs   = $matches['rhs']
							}
						}
					}
					
					# Pick a source rhs (same logic as you had, but without requiring old info to exist)
					$source = $null
					foreach ($e in $entries)
					{
						if ($e.Key.Length -ge 3 -and $e.Key.Substring($e.Key.Length - 3) -eq $newTerminal) { $source = $e; break }
					}
					if (-not $source -and $entries.Count -gt 0) { $source = $entries[0] }
					
					$rhsUpdated = ""
					$wsToUse = ""
					if ($source)
					{
						$rhsUpdated = $source.Rhs
						$wsToUse = $source.Ws
					}
					
					# Replace any old processor keys inside RHS -> newKey
					foreach ($e in $entries)
					{
						$rhsUpdated = $rhsUpdated -replace ("(?<!\d)" + [regex]::Escape($e.Key) + "(?!\d)"), $newKey
					}
					
					# Strip existing REDIRMAIL/REDIRMSG then enforce <STORE>901
					$rhsUpdated = [regex]::Replace($rhsUpdated, '\bREDIRMAIL\s*=\s*\d{6,7}', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
					$rhsUpdated = [regex]::Replace($rhsUpdated, '\bREDIRMSG\s*=\s*\d{6,7}', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
					$rhsUpdated = $rhsUpdated -replace ',{2,}', ','
					$rhsUpdated = $rhsUpdated.Trim().Trim(',').Trim()
					
					if ([string]::IsNullOrWhiteSpace($rhsUpdated))
					{
						$rhsUpdated = "REDIRMAIL=$redirKey901,REDIRMSG=$redirKey901,TARGETSEND=,TARGETRECV=,DEADLOCKPRIORITY="
					}
					else
					{
						$rhsUpdated = "REDIRMAIL=$redirKey901,REDIRMSG=$redirKey901," + $rhsUpdated
					}
					
					$newProcessorLine = $wsToUse + $newKey + "=" + $rhsUpdated
					
					# Purge all numeric processor lines; insert ONE just after section header
					$out = @()
					for ($i = 0; $i -lt $lines.Length; $i++)
					{
						$out += $lines[$i]
						
						if ($i -eq $secStart)
						{
							$out += $newProcessorLine
							continue
						}
						
						if ($i -gt $secStart -and $i -lt $secEnd)
						{
							if ($lines[$i] -match '^\s*\d{6,7}\s*=') { $out = $out[0 .. ($out.Count - 2)]; continue } # remove what we just added
						}
					}
					
					$lines = $out
				}
			}
			
			# ---- FORCED write (always) ----
			[System.IO.File]::WriteAllText($SmsHttpsIniPath, ($lines -join $nl), $enc)
			Write-Host ("Updated SmsHttps.INI (forced): " + $SmsHttpsIniPath) -ForegroundColor Green
		}
		catch
		{
			$success = $false
			Write-Host "Failed updating SmsHttps.INI: $($_.Exception.Message)" -ForegroundColor Red
		}
	}
	return $success
}

# ===================================================================================================
#                       FUNCTION: Update_SQL_Tables_For_Machine_Name_Change
# ---------------------------------------------------------------------------------------------------
# Description:
#   Updates TER_TAB, STO_TAB, LNK_TAB, RUN_TAB, and STD_TAB after machine name change / store change.
#
# Auto DB context detection (based on Startup.ini DBNAME already gathered via Get_Database_Connection_String):
#   - LANESQL             => LocalTerminalDB (lane/POS/SCO cleanup in same store is allowed)
#   - STORESQL or HOSTSQL => ServerStoreDB  (do NOT lane-cleanup other terminals)
#
# STD_TAB enhancement:
#   - Also updates F1531 (CompanyName) store suffix when present (e.g., "Company Name 231" or "Company #0231")
#
# FIXES:
#   - Non-server runs preserve/copy TER_TAB 901 into the NEW store before deleting old-store rows.
#   - STO_TAB now always ensures the 901 row exists.
#   - LNK_TAB now preserves/copies 901 into the NEW store before cleanup, so the server side is not left empty.
# ===================================================================================================

function Update_SQL_Tables_For_Machine_Name_Change
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidatePattern('^(?!0+$)\d{3,4}$')]
		[string]$storeNumber,
		[Parameter(Mandatory = $true)]
		[string]$machineName,
		[Parameter(Mandatory = $true)]
		[string]$machineNumber,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineName,
		[Parameter(Mandatory = $false)]
		[string]$OldMachineNumber,
		[Parameter(Mandatory = $false)]
		[ValidateSet('Auto', 'LocalTerminalDB', 'ServerStoreDB')]
		[string]$DatabaseContext = 'Auto'
	)
	
	# ---------------------------
	# Constants / normalization
	# ---------------------------
	$hostStoreNumber = "999"
	$serverTerminal = "901"
	$protectedMinTerminal = 940
	
	$storeNumber = (($storeNumber | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	$machineName = (($machineName | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	$machineNumber = (($machineNumber | ForEach-Object { if ($null -eq $_) { "" }
				else { $_.ToString() } }).Trim())
	
	# Detect SERVER moniker (example: 0242SERVER001)
	$isServerMoniker = $false
	if (-not [string]::IsNullOrWhiteSpace($machineName) -and ($machineName -match '(?i)SERVER'))
	{
		$isServerMoniker = $true
	}
	elseif ($script:NodeRole -and ($script:NodeRole -match '^(?i)(StoreServer|HostServer)$'))
	{
		$isServerMoniker = $true
	}
	
	$null = Ensure_Database_Connection_Context
	
	# ---------------------------
	# Pull DBNAME (already gathered by Get_Database_Connection_String) for better Auto context
	# ---------------------------
	$dbNameDetected = $null
	try
	{
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('DBNAME') -and -not [string]::IsNullOrWhiteSpace([string]$script:FunctionResults['DBNAME']))
		{
			$dbNameDetected = ([string]$script:FunctionResults['DBNAME']).Trim()
		}
	}
	catch { }
	
	# Fallback: parse from ConnectionString if present
	if ([string]::IsNullOrWhiteSpace($dbNameDetected))
	{
		try
		{
			if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConnectionString'))
			{
				$cs = [string]$script:FunctionResults['ConnectionString']
				if ($cs -match '(?i)\bDatabase\s*=\s*([^;]+)') { $dbNameDetected = $Matches[1].Trim() }
				elseif ($cs -match '(?i)\bInitial\s+Catalog\s*=\s*([^;]+)') { $dbNameDetected = $Matches[1].Trim() }
			}
		}
		catch { }
	}
	
	# ---------------------------
	# Normalize NEW/OLD server host string (for matching TER_TAB.F1125 and replacing STO_TAB.F1018)
	# ---------------------------
	$newServerNameNorm = $machineName
	if ($null -eq $newServerNameNorm) { $newServerNameNorm = "" }
	$newServerNameNorm = $newServerNameNorm.Trim()
	$newServerNameNorm = $newServerNameNorm -replace '^[\\\/]+', ''
	if ($newServerNameNorm -match '[\\\/]') { $newServerNameNorm = ($newServerNameNorm -split '[\\\/]')[0] }
	if ($newServerNameNorm -match '\.') { $newServerNameNorm = ($newServerNameNorm -split '\.')[0] }
	$newServerNameNorm = $newServerNameNorm.Trim().ToUpper()
	
	$oldServerNameNorm = $OldMachineName
	if ([string]::IsNullOrWhiteSpace($oldServerNameNorm)) { $oldServerNameNorm = $env:COMPUTERNAME }
	if ($null -eq $oldServerNameNorm) { $oldServerNameNorm = "" }
	$oldServerNameNorm = $oldServerNameNorm.Trim()
	$oldServerNameNorm = $oldServerNameNorm -replace '^[\\\/]+', ''
	if ($oldServerNameNorm -match '[\\\/]') { $oldServerNameNorm = ($oldServerNameNorm -split '[\\\/]')[0] }
	if ($oldServerNameNorm -match '\.') { $oldServerNameNorm = ($oldServerNameNorm -split '\.')[0] }
	$oldServerNameNorm = $oldServerNameNorm.Trim().ToUpper()
	
	# ---------------------------
	# Normalize NEW terminal to 3 digits
	# ---------------------------
	$newTerminal = $null
	if ($machineNumber -match '^\d{1,3}$')
	{
		$newTerminal = ([int]$machineNumber).ToString("D3")
		if ($newTerminal -eq "000") { $newTerminal = $null }
	}
	else
	{
		$tmp = $machineNumber
		if ($null -eq $tmp) { $tmp = "" }
		$tmp = $tmp.Trim()
		$tmp = $tmp -replace '^[\\\/]+', ''
		if ($tmp -match '[\\\/]') { $tmp = ($tmp -split '[\\\/]')[0] }
		if ($tmp -match '\.') { $tmp = ($tmp -split '\.')[0] }
		$tmp = $tmp.Trim().ToUpper()
		
		if ($tmp -match '(\d{1,3})$')
		{
			$newTerminal = ([int]$Matches[1]).ToString("D3")
			if ($newTerminal -eq "000") { $newTerminal = $null }
		}
	}
	
	# Determine OLD terminal for safe RUN_TAB swap (used only when NOT server)
	if ([string]::IsNullOrWhiteSpace($OldMachineName)) { $OldMachineName = $env:COMPUTERNAME }
	
	$oldTerminal = $null
	
	if (-not [string]::IsNullOrWhiteSpace($OldMachineNumber))
	{
		if ($OldMachineNumber -match '^\d{1,3}$')
		{
			$oldTerminal = ([int]$OldMachineNumber).ToString("D3")
			if ($oldTerminal -eq "000") { $oldTerminal = $null }
		}
		else
		{
			$tmpOld = $OldMachineNumber
			if ($null -eq $tmpOld) { $tmpOld = "" }
			$tmpOld = $tmpOld.Trim()
			$tmpOld = $tmpOld -replace '^[\\\/]+', ''
			if ($tmpOld -match '[\\\/]') { $tmpOld = ($tmpOld -split '[\\\/]')[0] }
			if ($tmpOld -match '\.') { $tmpOld = ($tmpOld -split '\.')[0] }
			$tmpOld = $tmpOld.Trim().ToUpper()
			
			if ($tmpOld -match '(\d{1,3})$')
			{
				$oldTerminal = ([int]$Matches[1]).ToString("D3")
				if ($oldTerminal -eq "000") { $oldTerminal = $null }
			}
		}
	}
	
	if (-not $oldTerminal)
	{
		$tmpOld2 = $OldMachineName
		if ($null -eq $tmpOld2) { $tmpOld2 = "" }
		$tmpOld2 = $tmpOld2.Trim()
		$tmpOld2 = $tmpOld2 -replace '^[\\\/]+', ''
		if ($tmpOld2 -match '[\\\/]') { $tmpOld2 = ($tmpOld2 -split '[\\\/]')[0] }
		if ($tmpOld2 -match '\.') { $tmpOld2 = ($tmpOld2 -split '\.')[0] }
		$tmpOld2 = $tmpOld2.Trim().ToUpper()
		
		if ($tmpOld2 -match '(\d{1,3})$')
		{
			$oldTerminal = ([int]$Matches[1]).ToString("D3")
			if ($oldTerminal -eq "000") { $oldTerminal = $null }
		}
	}
	
	# ---------------------------
	# Decide DB context
	# ---------------------------
	$isLocalTerminalDB = $false
	
	if ($DatabaseContext -eq 'LocalTerminalDB')
	{
		$isLocalTerminalDB = $true
	}
	elseif ($DatabaseContext -eq 'ServerStoreDB')
	{
		$isLocalTerminalDB = $false
	}
	else
	{
		# Auto: prefer DBNAME
		if (-not [string]::IsNullOrWhiteSpace($dbNameDetected) -and ($dbNameDetected -match '^(?i)LANESQL$'))
		{
			$isLocalTerminalDB = $true
		}
		elseif (-not [string]::IsNullOrWhiteSpace($dbNameDetected) -and ($dbNameDetected -match '^(?i)(STORESQL|HOSTSQL)$'))
		{
			$isLocalTerminalDB = $false
		}
		else
		{
			# Fallback heuristic: if targeting this box, assume local terminal DB
			$execHostNorm = $env:COMPUTERNAME
			if ($null -eq $execHostNorm) { $execHostNorm = "" }
			$execHostNorm = ($execHostNorm + "").Trim()
			$execHostNorm = $execHostNorm -replace '^[\\\/]+', ''
			if ($execHostNorm -match '[\\\/]') { $execHostNorm = ($execHostNorm -split '[\\\/]')[0] }
			if ($execHostNorm -match '\.') { $execHostNorm = ($execHostNorm -split '\.')[0] }
			$execHostNorm = $execHostNorm.Trim().ToUpper()
			
			$targetHostNorm = $machineName
			if ($null -eq $targetHostNorm) { $targetHostNorm = "" }
			$targetHostNorm = ($targetHostNorm + "").Trim()
			$targetHostNorm = $targetHostNorm -replace '^[\\\/]+', ''
			if ($targetHostNorm -match '[\\\/]') { $targetHostNorm = ($targetHostNorm -split '[\\\/]')[0] }
			if ($targetHostNorm -match '\.') { $targetHostNorm = ($targetHostNorm -split '\.')[0] }
			$targetHostNorm = $targetHostNorm.Trim().ToUpper()
			
			$isLocalTerminalDB = ($targetHostNorm -eq $execHostNorm)
		}
	}
	
	# SAFETY: Non-server mode must have a terminal, otherwise we can't safely upsert/cleanup.
	if (-not $isServerMoniker -and -not $newTerminal)
	{
		return @{
			Success		    = $false
			FailedCommands  = @("Refusing to run: could not derive a valid NEW terminal from -machineNumber/-machineName for non-server mode.")
			IsServerMoniker = $false
			RunTabUpdated   = $false
			OldTerminalUsed = $oldTerminal
			NewTerminalUsed = $newTerminal
			DatabaseContext = $DatabaseContext
			LocalTerminalDB = $isLocalTerminalDB
			DBNAME		    = $dbNameDetected
		}
	}
	
	# ---------------------------
	# Table names
	# ---------------------------
	$terTableName = "TER_TAB"
	$stoTableName = "STO_TAB"
	$lnkTableName = "LNK_TAB"
	$runTableName = "RUN_TAB"
	$stdTableName = "STD_TAB"
	
	# ===============================================================================================
	# TER_TAB
	# ===============================================================================================
	
	$createViewCommandTer = @"
IF OBJECT_ID('Ter_Load', 'V') IS NOT NULL DROP VIEW Ter_Load;
EXEC(N'
CREATE VIEW Ter_Load AS
SELECT F1056, F1057, F1058, F1125, F1169
FROM $terTableName;
');
"@
	
	$moveProtectedTer_ToNewStore = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @ProtectedMin int = $protectedMinTerminal;

INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
SELECT @NewStore, t.F1057, t.F1058, t.F1125, t.F1169
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) >= @ProtectedMin
  AND NOT EXISTS (
      SELECT 1
      FROM $terTableName x
      WHERE x.F1056 = @NewStore
        AND x.F1057 = t.F1057
  );

DELETE t
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) >= @ProtectedMin;
"@
	
	$moveServerNodes902_920_ToNewStore_AndUpdate = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

;WITH Candidates AS
(
    SELECT t.*
    FROM $terTableName t
    WHERE t.F1056 NOT IN (@NewStore, @HostStore)
      AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
      AND t.F1125 IS NOT NULL
      AND
      (
          UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
      )
)
INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
SELECT @NewStore, c.F1057, c.F1058, c.F1125, c.F1169
FROM Candidates c
WHERE NOT EXISTS (
    SELECT 1
    FROM $terTableName x
    WHERE x.F1056 = @NewStore
      AND x.F1057 = c.F1057
);

UPDATE t
SET
    t.F1125 = CASE
                WHEN t.F1125 IS NULL THEN t.F1125
                WHEN LEFT(LTRIM(RTRIM(t.F1125)),1)='@' THEN '@' + @NewServer
                ELSE REPLACE(UPPER(t.F1125), @OldServer, @NewServer)
              END,
    t.F1169 = CASE
                WHEN t.F1169 IS NULL THEN t.F1169
                ELSE '\\' + @NewServer + '\storeman\office\XF' + @NewStore + '901\'
              END
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920;

DELETE t
FROM $terTableName t
WHERE t.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
  AND t.F1125 IS NOT NULL
  AND
  (
      UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
      OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
      OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
      OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
  );
"@
	
	$deleteOtherStoresTer = @"
DELETE FROM $terTableName
WHERE F1056 NOT IN ('$storeNumber', '$hostStoreNumber');
"@
	
	$deleteNonServerRowsForNewStore = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;
DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

DELETE t
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND t.F1057 <> @ServerTerm
  AND
  (
      (
          TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) IS NULL
          OR TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) < @ProtectedMin
      )
      AND
      (
          NOT (
              TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
              AND t.F1125 IS NOT NULL
              AND
              (
                  UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
                  OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
                  OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
                  OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
              )
          )
      )
  );
"@
	
	$upsertLaneTer = @"
IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056='$storeNumber' AND F1057='$newTerminal')
BEGIN
    UPDATE $terTableName
    SET F1058='Terminal $newTerminal',
        F1125='\\$machineName\storeman\office\XF$storeNumber$newTerminal\',
        F1169='\\$machineName\storeman\office\XF$storeNumber$serverTerminal\'
    WHERE F1056='$storeNumber' AND F1057='$newTerminal';
END
ELSE
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169) VALUES
    ('$storeNumber', '$newTerminal',
     'Terminal $newTerminal',
     '\\$machineName\storeman\office\XF$storeNumber$newTerminal\',
     '\\$machineName\storeman\office\XF$storeNumber$serverTerminal\');
END
"@
	
	$preserve901ForNonServerStoreChange = @"
DECLARE @NewStore   varchar(10) = '$storeNumber';
DECLARE @HostStore  varchar(10) = '$hostStoreNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';

IF NOT EXISTS (
    SELECT 1
    FROM $terTableName
    WHERE F1056 = @NewStore
      AND F1057 = @ServerTerm
)
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169)
    SELECT TOP 1
           @NewStore,
           @ServerTerm,
           CASE
               WHEN NULLIF(LTRIM(RTRIM(t.F1058)), '') IS NULL THEN 'Server'
               ELSE t.F1058
           END,
           t.F1125,
           t.F1169
    FROM $terTableName t
    WHERE t.F1057 = @ServerTerm
      AND t.F1056 <> @HostStore
    ORDER BY
        CASE WHEN t.F1125 IS NOT NULL AND LTRIM(RTRIM(t.F1125)) <> '' THEN 0 ELSE 1 END,
        t.F1056;
END

UPDATE $terTableName
SET F1058 = CASE
                WHEN NULLIF(LTRIM(RTRIM(F1058)), '') IS NULL THEN 'Server'
                ELSE F1058
            END
WHERE F1056 = @NewStore
  AND F1057 = @ServerTerm;
"@
	
	$cleanupTer_ForLocalTerminalDB = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @KeepTerm varchar(10) = '$newTerminal';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DELETE t
FROM $terTableName t
WHERE t.F1056 = @NewStore
  AND t.F1057 NOT IN (@KeepTerm, @ServerTerm)
  AND (TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) IS NULL OR TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) < @ProtectedMin);
"@
	
	$upsertServer901Ter_ForNewStoreOnly = @"
IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056='$storeNumber' AND F1057='$serverTerminal')
BEGIN
    UPDATE $terTableName
    SET F1058='Server',
        F1125='@$machineName',
        F1169=NULL
    WHERE F1056='$storeNumber' AND F1057='$serverTerminal';
END
ELSE
BEGIN
    INSERT INTO $terTableName (F1056, F1057, F1058, F1125, F1169) VALUES
    ('$storeNumber', '$serverTerminal',
     'Server',
     '@$machineName',
     NULL);
END
"@
	
	$dropViewCommandTer = @"
IF OBJECT_ID('Ter_Load', 'V') IS NOT NULL DROP VIEW Ter_Load;
"@
	
	# ===============================================================================================
	# STO_TAB
	# ===============================================================================================
	
	$createViewCommandSto = @"
IF OBJECT_ID('Sto_Load', 'V') IS NOT NULL DROP VIEW Sto_Load;
EXEC(N'
CREATE VIEW Sto_Load AS
SELECT F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691
FROM $stoTableName;
');
"@
	
	$upsertStoTerminal = @"
MERGE INTO $stoTableName AS target
USING (VALUES
    ('$newTerminal', 'Terminal $newTerminal', 1, 1, 1, NULL, NULL, NULL, NULL)
) AS source (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
ON target.F1000 = source.F1000
WHEN MATCHED THEN
    UPDATE SET
        F1018 = source.F1018,
        F1180 = source.F1180,
        F1181 = source.F1181,
        F1182 = source.F1182
WHEN NOT MATCHED THEN
    INSERT (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
    VALUES (source.F1000, source.F1018, source.F1180, source.F1181, source.F1182, source.F1937, source.F1965, source.F1966, source.F2691);
"@
	
	$upsertStoServer901 = @"
MERGE INTO $stoTableName AS target
USING (VALUES
    ('$serverTerminal', 'Server', 1, 1, 1, NULL, NULL, NULL, NULL)
) AS source (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
ON target.F1000 = source.F1000
WHEN MATCHED THEN
    UPDATE SET
        F1018 = CASE
                    WHEN NULLIF(LTRIM(RTRIM(target.F1018)), '') IS NULL THEN source.F1018
                    ELSE target.F1018
                END,
        F1180 = ISNULL(target.F1180, source.F1180),
        F1181 = ISNULL(target.F1181, source.F1181),
        F1182 = ISNULL(target.F1182, source.F1182)
WHEN NOT MATCHED THEN
    INSERT (F1000, F1018, F1180, F1181, F1182, F1937, F1965, F1966, F2691)
    VALUES (source.F1000, source.F1018, source.F1180, source.F1181, source.F1182, source.F1937, source.F1965, source.F1966, source.F2691);
"@
	
	$cleanupSto_ForLocalTerminalDB = @"
DECLARE @KeepTerm varchar(10) = '$newTerminal';

DELETE FROM $stoTableName
WHERE F1000 <> @KeepTerm
  AND F1000 NOT IN ('DSM','PAL','RAL','XAL','901','999');
"@
	
	$cleanupSto_ForServer = @"
;WITH StoNums AS
(
    SELECT
        F1000,
        F1018,
        TRY_CONVERT(int, LTRIM(RTRIM(F1000))) AS NumVal
    FROM $stoTableName
    WHERE ISNUMERIC(F1000) = 1
)
DELETE s
FROM $stoTableName s
JOIN StoNums n ON n.F1000 = s.F1000
WHERE
    n.F1000 NOT IN ('901','999')
    AND
    (
        (n.NumVal BETWEEN 1 AND 99)
        OR
        (
            n.NumVal BETWEEN 902 AND 920
            AND (n.F1018 IS NULL OR UPPER(n.F1018) NOT LIKE '%SERVER%')
        )
    );

DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

UPDATE $stoTableName
SET F1018 = REPLACE(UPPER(F1018), @OldServer, @NewServer)
WHERE ISNUMERIC(F1000) = 1
  AND TRY_CONVERT(int, LTRIM(RTRIM(F1000))) BETWEEN 902 AND 920
  AND F1018 IS NOT NULL
  AND UPPER(F1018) LIKE '%' + @OldServer + '%';
"@
	
	$dropViewCommandSto = @"
IF OBJECT_ID('Sto_Load', 'V') IS NOT NULL DROP VIEW Sto_Load;
"@
	
	# ===============================================================================================
	# LNK_TAB
	# ===============================================================================================
	
	$createViewCommandLnk = @"
IF OBJECT_ID('Lnk_Load', 'V') IS NOT NULL DROP VIEW Lnk_Load;
EXEC(N'
CREATE VIEW Lnk_Load AS
SELECT F1000, F1056, F1057
FROM $lnkTableName;
');
"@
	
	$preserve901LnkForNewStore = @"
DECLARE @NewStore   varchar(10) = '$storeNumber';
DECLARE @HostStore  varchar(10) = '$hostStoreNumber';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT DISTINCT
       l.F1000,
       @NewStore,
       @ServerTerm
FROM $lnkTableName l
WHERE l.F1057 = @ServerTerm
  AND l.F1056 <> @HostStore
  AND l.F1000 IS NOT NULL
  AND NOT EXISTS (
      SELECT 1
      FROM $lnkTableName x
      WHERE x.F1000 = l.F1000
        AND x.F1056 = @NewStore
        AND x.F1057 = @ServerTerm
  );

IF NOT EXISTS (
    SELECT 1
    FROM $lnkTableName
    WHERE F1056 = @NewStore
      AND F1057 = @ServerTerm
)
BEGIN
    INSERT INTO $lnkTableName (F1000, F1056, F1057)
    VALUES (@ServerTerm, @NewStore, @ServerTerm);
END
"@
	
	$rebuildLnkForTerminal = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';
DECLARE @Term      varchar(10) = '$newTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DECLARE @NewServer varchar(200) = '$newServerNameNorm';
DECLARE @OldServer varchar(200) = '$oldServerNameNorm';

DECLARE @SourceStore varchar(10);

SELECT TOP 1 @SourceStore = F1056
FROM $lnkTableName
WHERE F1057 = @Term
  AND F1056 <> @HostStore
  AND F1056 <> @NewStore
ORDER BY F1056;

IF @SourceStore IS NULL
BEGIN
    SELECT TOP 1 @SourceStore = F1056
    FROM $lnkTableName
    WHERE F1057 = @Term
      AND F1056 = @NewStore;
END

DECLARE @Codes TABLE (F1000 varchar(50) NOT NULL PRIMARY KEY);

IF @SourceStore IS NOT NULL
BEGIN
    INSERT INTO @Codes(F1000)
    SELECT DISTINCT F1000
    FROM $lnkTableName
    WHERE F1056 = @SourceStore
      AND F1057 = @Term
      AND F1000 IS NOT NULL;
END

IF EXISTS (SELECT 1 FROM $terTableName WHERE F1056=@NewStore AND TRY_CONVERT(int, LTRIM(RTRIM(F1057))) BETWEEN 902 AND 920)
BEGIN
    DECLARE @ServerTerms TABLE (Term varchar(10) NOT NULL PRIMARY KEY);

    INSERT INTO @ServerTerms(Term)
    SELECT DISTINCT t.F1057
    FROM $terTableName t
    WHERE t.F1056 = @NewStore
      AND TRY_CONVERT(int, LTRIM(RTRIM(t.F1057))) BETWEEN 902 AND 920
      AND t.F1125 IS NOT NULL
      AND
      (
          UPPER(t.F1125) LIKE '%\\' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%\\' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%@' + @OldServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @NewServer + '%'
          OR UPPER(t.F1125) LIKE '%' + @OldServer + '%'
      );

    INSERT INTO $lnkTableName (F1000, F1056, F1057)
    SELECT l.F1000, @NewStore, l.F1057
    FROM $lnkTableName l
    JOIN @ServerTerms st ON st.Term = l.F1057
    WHERE l.F1056 NOT IN (@NewStore, @HostStore)
      AND NOT EXISTS (
          SELECT 1
          FROM $lnkTableName x
          WHERE x.F1000 = l.F1000
            AND x.F1056 = @NewStore
            AND x.F1057 = l.F1057
      );

    DELETE l
    FROM $lnkTableName l
    JOIN @ServerTerms st ON st.Term = l.F1057
    WHERE l.F1056 NOT IN (@NewStore, @HostStore);
END

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT l.F1000, @NewStore, l.F1057
FROM $lnkTableName l
WHERE l.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) >= @ProtectedMin
  AND NOT EXISTS (
      SELECT 1
      FROM $lnkTableName x
      WHERE x.F1000 = l.F1000
        AND x.F1056 = @NewStore
        AND x.F1057 = l.F1057
  );

DELETE l
FROM $lnkTableName l
WHERE l.F1056 NOT IN (@NewStore, @HostStore)
  AND TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) >= @ProtectedMin;

DELETE FROM $lnkTableName
WHERE F1056 NOT IN (@NewStore, @HostStore);

IF NOT EXISTS (SELECT 1 FROM @Codes)
BEGIN
    INSERT INTO @Codes(F1000) VALUES (@Term);
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='DSM') INSERT INTO @Codes(F1000) VALUES ('DSM');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='PAL') INSERT INTO @Codes(F1000) VALUES ('PAL');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='RAL') INSERT INTO @Codes(F1000) VALUES ('RAL');
    IF NOT EXISTS (SELECT 1 FROM @Codes WHERE F1000='XAL') INSERT INTO @Codes(F1000) VALUES ('XAL');
END

INSERT INTO $lnkTableName (F1000, F1056, F1057)
SELECT c.F1000, @NewStore, @Term
FROM @Codes c
WHERE NOT EXISTS (
    SELECT 1
    FROM $lnkTableName x
    WHERE x.F1000 = c.F1000
      AND x.F1056 = @NewStore
      AND x.F1057 = @Term
);
"@
	
	$cleanupLnk_ForLocalTerminalDB = @"
DECLARE @NewStore varchar(10) = '$storeNumber';
DECLARE @KeepTerm varchar(10) = '$newTerminal';
DECLARE @ServerTerm varchar(10) = '$serverTerminal';
DECLARE @ProtectedMin int = $protectedMinTerminal;

DELETE l
FROM $lnkTableName l
WHERE l.F1056 = @NewStore
  AND l.F1057 NOT IN (@KeepTerm, @ServerTerm)
  AND (TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) IS NULL OR TRY_CONVERT(int, LTRIM(RTRIM(l.F1057))) < @ProtectedMin);
"@
	
	$dropViewCommandLnk = @"
IF OBJECT_ID('Lnk_Load', 'V') IS NOT NULL DROP VIEW Lnk_Load;
"@
	
	# ===============================================================================================
	# RUN_TAB (ONLY when NOT server)
	# ===============================================================================================
	
	$createViewCommandRun = @"
IF OBJECT_ID('Run_Load', 'V') IS NOT NULL DROP VIEW Run_Load;
EXEC(N'
CREATE VIEW Run_Load AS
SELECT F1102, F1000, F1104
FROM $runTableName;
');
"@
	
	$updateRunTab_SafeSwap_OldToNew = @"
DECLARE @OldTerm varchar(10) = '$oldTerminal';
DECLARE @NewTerm varchar(10) = '$newTerminal';
DECLARE @Server  varchar(10) = '$serverTerminal';

IF @OldTerm IS NOT NULL AND @OldTerm <> '' AND @OldTerm <> @NewTerm
BEGIN
    UPDATE $runTableName
    SET F1000 = @NewTerm
    WHERE F1000 = @OldTerm;

    UPDATE $runTableName
    SET F1104 = @NewTerm
    WHERE F1104 = @OldTerm
      AND F1104 <> @Server;
END
"@
	
	$dropViewCommandRun = @"
IF OBJECT_ID('Run_Load', 'V') IS NOT NULL DROP VIEW Run_Load;
"@
	
	# ===============================================================================================
	# STD_TAB (store key + CompanyName suffix fix in F1531)
	# ===============================================================================================
	
	$createViewCommandStd = @"
IF OBJECT_ID('Std_Load', 'V') IS NOT NULL DROP VIEW Std_Load;
EXEC(N'
CREATE VIEW Std_Load AS
SELECT F1056, F1531
FROM $stdTableName;
');
"@
	
	$fixStdTabStoreKey_AndCompanySuffix = @"
DECLARE @NewStore  varchar(10) = '$storeNumber';
DECLARE @HostStore varchar(10) = '$hostStoreNumber';

-- =========================================
-- 1) Force store key: keep ONLY (@NewStore + @HostStore)
--    - If @NewStore row exists: delete all other non-host rows
--    - Else: convert ONE non-host row to @NewStore, then delete the rest
-- =========================================
IF EXISTS (SELECT 1 FROM $stdTableName WHERE F1056 = @NewStore)
BEGIN
    DELETE FROM $stdTableName
    WHERE F1056 <> @NewStore
      AND F1056 <> @HostStore;
END
ELSE
BEGIN
    DECLARE @PickOld varchar(10) = NULL;

    SELECT TOP 1 @PickOld = F1056
    FROM $stdTableName
    WHERE F1056 <> @HostStore
    ORDER BY F1056;

    IF @PickOld IS NOT NULL
    BEGIN
        UPDATE $stdTableName
        SET F1056 = @NewStore
        WHERE F1056 = @PickOld;
    END

    DELETE FROM $stdTableName
    WHERE F1056 <> @NewStore
      AND F1056 <> @HostStore;
END

-- =========================================
-- 2) Update F1531 ONLY IF a store number is already present at the end
--    (space/#/- + 3 or 4 digits). If no suffix exists -> do nothing.
--    Robust against NBSP/TAB/CR/LF at the end.
-- =========================================
;WITH X AS
(
    SELECT
        s.F1056,
        s.F1531,
        Clean = RTRIM(
                    REPLACE(REPLACE(REPLACE(REPLACE(
                        LTRIM(RTRIM(s.F1531)),
                        CHAR(160), ' '),
                        CHAR(9),  ' '),
                        CHAR(13), ' '),
                        CHAR(10), ' ')
                )
    FROM STD_TAB s
    WHERE s.F1056 = @NewStore
      AND s.F1531 IS NOT NULL
)
UPDATE X
SET F1531 =
    CASE
        WHEN Clean LIKE '%[ #\-][0-9][0-9][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 4)) IS NOT NULL
             AND RIGHT(Clean, 4) NOT IN ('0901','0999')
        THEN LEFT(Clean, LEN(Clean) - 4) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 3)) IS NOT NULL
             AND RIGHT(Clean, 3) NOT IN ('901','999')
        THEN LEFT(Clean, LEN(Clean) - 3) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 2)) IS NOT NULL
        THEN LEFT(Clean, LEN(Clean) - 2) + @NewStore

        WHEN Clean LIKE '%[ #\-][0-9]'
             AND TRY_CONVERT(int, RIGHT(Clean, 1)) IS NOT NULL
        THEN LEFT(Clean, LEN(Clean) - 1) + @NewStore

        ELSE F1531
    END
WHERE
    Clean LIKE '%[ #\-][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9][0-9]'
    OR Clean LIKE '%[ #\-][0-9][0-9][0-9][0-9]';
"@
	
	$dropViewCommandStd = @"
IF OBJECT_ID('Std_Load', 'V') IS NOT NULL DROP VIEW Std_Load;
"@
	
	# ===============================================================================================
	# Execute SQL commands (ORDER MATTERS)
	# ===============================================================================================
	
	$sqlCommands = @(
		$createViewCommandTer,
		$moveProtectedTer_ToNewStore
	)
	
	if ($isServerMoniker)
	{
		$sqlCommands += $moveServerNodes902_920_ToNewStore_AndUpdate
		$sqlCommands += $deleteOtherStoresTer
		$sqlCommands += $deleteNonServerRowsForNewStore
		$sqlCommands += $upsertServer901Ter_ForNewStoreOnly
	}
	else
	{
		$sqlCommands += $preserve901ForNonServerStoreChange
		$sqlCommands += $deleteOtherStoresTer
		$sqlCommands += $upsertLaneTer
		
		if ($isLocalTerminalDB)
		{
			$sqlCommands += $cleanupTer_ForLocalTerminalDB
		}
	}
	
	$sqlCommands += @(
		$dropViewCommandTer,
		
		# STO_TAB
		$createViewCommandSto,
		$upsertStoServer901
	)
	
	if ($isServerMoniker)
	{
		$sqlCommands += $cleanupSto_ForServer
	}
	else
	{
		$sqlCommands += $upsertStoTerminal
		if ($isLocalTerminalDB)
		{
			$sqlCommands += $cleanupSto_ForLocalTerminalDB
		}
	}
	
	$sqlCommands += @(
		$dropViewCommandSto,
		
		# LNK_TAB
		$createViewCommandLnk,
		$preserve901LnkForNewStore,
		$rebuildLnkForTerminal
	)
	
	if (-not $isServerMoniker -and $isLocalTerminalDB)
	{
		$sqlCommands += $cleanupLnk_ForLocalTerminalDB
	}
	
	$sqlCommands += @(
		$dropViewCommandLnk
	)
	
	if (-not $isServerMoniker)
	{
		$sqlCommands += @(
			$createViewCommandRun,
			$updateRunTab_SafeSwap_OldToNew,
			$dropViewCommandRun
		)
	}
	
	$sqlCommands += @(
		$createViewCommandStd,
		$fixStdTabStoreKey_AndCompanySuffix,
		$dropViewCommandStd
	)
	
	$allSqlSuccessful = $true
	$failedSqlCommands = @()
	
	foreach ($command in $sqlCommands)
	{
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$allSqlSuccessful = $false
			$failureMessage = if (-not [string]::IsNullOrWhiteSpace($script:LastSqlExecutionError)) { $script:LastSqlExecutionError } else { "Unknown SQL execution failure." }
			$failedSqlCommands += ("Error: " + $failureMessage + "`r`nSQL:`r`n" + $command)
		}
	}
	
	return @{
		Success		    = $allSqlSuccessful
		FailedCommands  = $failedSqlCommands
		IsServerMoniker = $isServerMoniker
		RunTabUpdated   = (-not $isServerMoniker)
		OldTerminalUsed = $oldTerminal
		NewTerminalUsed = $newTerminal
		DatabaseContext = $DatabaseContext
		LocalTerminalDB = $isLocalTerminalDB
		DBNAME		    = $dbNameDetected
		ConfiguredDBSERVER = if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('ConfiguredDBSERVER')) { [string]$script:FunctionResults['ConfiguredDBSERVER'] } else { $null }
		RuntimeDBSERVER = if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('RuntimeDBSERVER')) { [string]$script:FunctionResults['RuntimeDBSERVER'] } else { $null }
	}
}

# ===================================================================================================
#                                       FUNCTION: RPT_Blank
# ---------------------------------------------------------------------------------------------------
# Description:
#   Truncates the specified list of tables in the SQL database.
# ===================================================================================================

function RPT_Blank
{
	param (
		[string[]]$tables
	)
	
	# Initialize an array to store failed truncate commands
	$failedTruncateTables = @()
	
	# Prepare and execute Truncate Commands
	foreach ($table in $tables)
	{
		$command = "TRUNCATE TABLE $table;"
		if (-not (Execute_SQL_Commands -commandText $command))
		{
			$failedTruncateTables += $table # Add failed table to the array
		}
	}
	
	# Return the list of failed tables
	return $failedTruncateTables
}

# ===================================================================================================
#                                       FUNCTION: Remove_GT_Registry_Values
# ---------------------------------------------------------------------------------------------------
# Description:
#   Removes all registry values starting with 'GT' from specified registry paths.
# ===================================================================================================

function Remove_GT_Registry_Values
{
	# Define registry paths for 32-bit and 64-bit
	$regPath32 = "HKLM:\SOFTWARE\Store Management\Counters"
	$regPath64 = "HKLM:\SOFTWARE\Wow6432Node\Store Management\Counters"
	
	$is64bit = [System.Environment]::Is64BitOperatingSystem
	
	$totalDeletedCount = 0
	$success = $true
	$status = "Successful"
	$message = ""
	
	# Build list of candidate paths (prefer Wow6432Node on 64-bit, but include both if present)
	$pathsToCheck = @()
	
	if ($is64bit)
	{
		if (Test-Path $regPath64) { $pathsToCheck += $regPath64 }
		if (Test-Path $regPath32) { $pathsToCheck += $regPath32 }
	}
	else
	{
		if (Test-Path $regPath32) { $pathsToCheck += $regPath32 }
	}
	
	if (-not $pathsToCheck -or $pathsToCheck.Count -eq 0)
	{
		return @{
			Success	     = $false
			Status	     = "Failed"
			DeletedCount = 0
			Message	     = "No valid registry paths found for the current environment."
		}
	}
	
	foreach ($path in $pathsToCheck)
	{
		try
		{
			# Get value names from the key by inspecting item properties
			# (Get-ItemProperty returns a PSCustomObject with properties = value names + PS metadata)
			$itemProps = Get-ItemProperty -Path $path -ErrorAction Stop
			
			# Collect only real registry value names starting with GT (exclude PS metadata props)
			$gtNames = @()
			foreach ($p in $itemProps.PSObject.Properties)
			{
				if ($p.Name -eq 'PSPath') { continue }
				if ($p.Name -eq 'PSParentPath') { continue }
				if ($p.Name -eq 'PSChildName') { continue }
				if ($p.Name -eq 'PSDrive') { continue }
				if ($p.Name -eq 'PSProvider') { continue }
				
				if ($p.Name -like 'GT*')
				{
					$gtNames += $p.Name
				}
			}
			
			# Delete each GT* value
			foreach ($name in $gtNames)
			{
				try
				{
					Remove-ItemProperty -Path $path -Name $name -ErrorAction Stop
					$totalDeletedCount++
				}
				catch
				{
					# Keep going but mark overall as partial failure
					$success = $false
					$status = "Failed"
					if ($message) { $message += " | " }
					$message += "Failed to delete '$name' in '$path': $_"
				}
			}
		}
		catch
		{
			$success = $false
			$status = "Failed"
			if ($message) { $message += " | " }
			$message += "Error accessing registry path '$path': $_"
		}
	}
	
	return @{
		Success	     = $success
		Status	     = $status
		DeletedCount = $totalDeletedCount
		Message	     = $message
	}
}

# ===================================================================================================
#                                  SECTION: Mini Ghost tool (workflows)
# ---------------------------------------------------------------------------------------------------
#   Lane-only tool. Two actions:
#     Run Mini-Ghost             prompts store # / lane # / IP, then: store number (if it changes) ->
#                                machine name + INI/SQL sync -> truncate -> GT registry -> IP -> reboot?
#     Update Store SQL Database  only the SQL sync (TER/STO/LNK/RUN/STD) + truncate, for a database
#                                copied from another lane (no rename, no INI, no IP)
#   The workflows run in a background runspace together with the Mini Ghost core functions; their
#   Write-Host output is streamed into the terminal (Update_MgJobOutput).
# ===================================================================================================

# Same list as Mini Ghost's "Truncate Tables" button
$script:MgTruncateTables = @(
	"TRS_LOG", "TRS_FIN", "TRS_DPT", "TRS_SUB", "TRS_CLK",
	"TRS_CLT", "TRS_VND", "TRS_ITM",
	"RPT_CLT_D", "RPT_CLT_W", "RPT_CLT_M", "RPT_CLT_N", "RPT_CLT_Y", "RPT_CLT_P", "RPT_CLT_F", "RPT_CLT_ITM_D", "RPT_CLT_ITM_N",
	"RPT_FIN", "RPT_DPT", "RPT_SUB", "RPT_HOU", "RPT_VND",
	"COST_REV", "POS_REV", "OBJ_REV", "PRICE_REV", "REV_HDR", "SAL_REG_SAV", "SAL_HDR_SAV", "SAL_TTL_SAV", "SAL_DET_SAV",
	"RPT_CLK_D", "RPT_CLK_W", "RPT_CLK_M", "RPT_CLK_Y", "RPT_CLK_P", "RPT_CLK_F", "RPT_CLK_N",
	"RPT_ITM_D", "RPT_ITM_W", "RPT_ITM_M", "RPT_ITM_Y", "RPT_ITM_P", "RPT_ITM_F", "RPT_ITM_N",
	"DATA_REG",
	"RPT_CHK", "RENT_TAB", "TIM_TAB", "GAS_INVENT", "GAS_COUNT", "GAS_TRANS",
	"SAL_BAT", "SAL_HDR", "SAL_REG", "SAL_DET", "SAL_TTL",
	"REC_BAT", "REC_HDR", "REC_REG", "REC_TTL", "INV_HDR", "INV_REG", "INV_TTL",
	"RPT_PAY_N", "RPT_PAY_M", "TRS_PAY", "OFR_TAB"
)

# Functions shipped into the Mini Ghost runspace
$script:MgRunspaceFunctions = @(
	'Set_Operation_Status', 'Set_Label_Text', 'Get_Effective_Machine_Name', 'Get_Database_Connection_String',
	'Ensure_Database_Connection_Context', 'Get_Store_Number_From_INI', 'Get_Store_Name_From_INI', 'Get_Active_IP_Config',
	'Remove_Old_XZ_Folders', 'Execute_SQL_Commands', 'Update_SQL_Tables_For_Store_Number_Change', 'Resolve_Startup_Ini_Path',
	'Get_Machine_Name_Context', 'Update_Startup_Ini_For_Machine', 'Invoke_Machine_Name_Sync', 'Update_INIs',
	'Update_SQL_Tables_For_Machine_Name_Change', 'RPT_Blank', 'Remove_GT_Registry_Values',
	'Initialize_MiniGhost_Context', 'Write_MgLog', 'Write_MgSummary', 'Set_MiniGhost_Network',
	'Invoke_MiniGhost_Full', 'Invoke_MiniGhost_DbUpdate'
)

$script:MgJob = $null
$script:MgInfo = $null

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Initialize_MiniGhost_Context
#   Sets the script-scope variables the Mini Ghost core reads (same as Mini Ghost's startup section),
#   from the storeman folder the toolkit already detected (no drive scan).
# ---------------------------------------------------------------------------------------------------
function Initialize_MiniGhost_Context
{
	param ([string]$StoremanRoot)
	if (-not $StoremanRoot) { $StoremanRoot = 'C:\storeman' }
	$script:FunctionResults = @{ }
	$script:newStoreNumber = $null
	$script:newMachineName = $null
	$script:ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)
	$script:utf8NoBOM = New-Object System.Text.UTF8Encoding($false)
	$script:currentMachineName = $env:COMPUTERNAME
	$script:BasePath = $StoremanRoot
	$script:OfficePath = Join-Path $StoremanRoot 'Office'
	$script:LoadPath = Join-Path $script:OfficePath 'Load'
	$script:StartupIniPath = Join-Path $StoremanRoot 'Startup.ini'
	$script:GlobalSmsStartIniPath = Join-Path $StoremanRoot 'SMSStart.ini'
	$script:SystemIniPath = Join-Path $script:OfficePath 'system.ini'
	$script:GasInboxPath = Join-Path $script:OfficePath 'XchGAS\INBOX'
	$script:DbsPath = Join-Path $script:OfficePath 'Dbs'
	$script:TempDir = [System.IO.Path]::GetTempPath()
	$script:SmsHttpsIniPath = $null
	foreach ($rel in @('SmsHttps64\SmsHttps.INI', 'SmsHttps32\SmsHttps.ini', 'SmsHttps\SmsHttps.ini'))
	{
		$p = Join-Path $StoremanRoot $rel
		if (Test-Path -LiteralPath $p) { $script:SmsHttpsIniPath = $p; break }
	}
	$script:WinIniPath = Get-ChildItem -LiteralPath $script:DbsPath -Filter 'INFO_*901_WIN.INI' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
	$script:SmsStartIniPath = Get-ChildItem -LiteralPath $script:DbsPath -Filter 'INFO_*901_SMSStart.ini' -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
	$script:operationStatus = @{ }
	foreach ($k in @('StoreNumberChange', 'MachineNameChange', 'OldXFoldersDeletion', 'StartupIniUpdate', 'IPConfiguration', 'TableTruncation', 'RegistryCleanup', 'SQLDatabaseUpdate'))
	{
		$script:operationStatus[$k] = @{ Status = 'Pending'; Message = ''; Details = '' }
	}
	# Node role from DBNAME, as Mini Ghost does
	$script:NodeRole = 'Unknown'
	try
	{
		$line = Get-Content -LiteralPath $script:StartupIniPath -ErrorAction Stop | Where-Object { $_ -match '^\s*(?i:DBNAME)\s*=' } | Select-Object -First 1
		if ($line -match '^\s*(?i:DBNAME)\s*=\s*(.+?)\s*$')
		{
			$dbn = $Matches[1].Trim().ToUpper()
			$script:FunctionResults['DBNAME'] = $dbn
			$script:NodeRole = switch -regex ($dbn) { '^LANESQL$' { 'Lane' } '^STORESQL$' { 'StoreServer' } '^HOSTSQL$' { 'HostServer' } default { 'Unknown' } }
		}
	}
	catch { }
}

# Coloured log line from the runspace (colour = terminal kind, see Update_MgJobOutput)
function Write_MgLog
{
	param ([string]$Text, [string]$Kind = 'text')
	$c = switch ($Kind) { 'ok' { 'Green' } 'fail' { 'Red' } 'warn' { 'Yellow' } 'info' { 'Cyan' } 'step' { 'Magenta' } 'muted' { 'DarkGray' } default { 'White' } }
	Write-Host $Text -ForegroundColor $c
}

# Mini Ghost "Show Summary", as terminal lines
function Write_MgSummary
{
	Write_MgLog '---- Summary ----' 'step'
	foreach ($k in @('StoreNumberChange', 'MachineNameChange', 'OldXFoldersDeletion', 'StartupIniUpdate', 'SQLDatabaseUpdate', 'TableTruncation', 'RegistryCleanup', 'IPConfiguration'))
	{
		$o = $script:operationStatus[$k]
		if (-not $o -or $o.Status -eq 'Pending') { continue }
		$kind = switch ($o.Status) { 'Successful' { 'ok' } 'Failed' { 'fail' } default { 'muted' } }
		$text = '{0}: {1}' -f $k, $o.Status
		if ($o.Message) { $text += ' - ' + $o.Message }
		if ($o.Details) { $text += '  (' + $o.Details + ')' }
		Write_MgLog $text $kind
	}
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Set_MiniGhost_Network - Mini Ghost "Configure Network" without the dialogs
#   $Net = @{ Mode = Keep | Dhcp | Static; Adapter; Ip; Gateway }   (static: /24, DNS 8.8.8.8 + 8.8.4.4)
# ---------------------------------------------------------------------------------------------------
function Set_MiniGhost_Network
{
	param ($Net)
	if (-not $Net -or $Net.Mode -eq 'Keep')
	{
		Write_MgLog 'IP address left unchanged.' 'muted'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'IPConfiguration' -Status 'Skipped' -Message 'IP left unchanged.'
		return
	}
	$adapterName = $Net.Adapter
	if ($Net.Mode -eq 'Dhcp')
	{
		$out = netsh interface ip set address name="$adapterName" source=dhcp 2>&1
		$code = $LASTEXITCODE
		$null = netsh interface ip set dns name="$adapterName" source=dhcp 2>&1
		if ($code -eq 0)
		{
			Write_MgLog "DHCP configuration applied on '$adapterName'." 'ok'
			Set_Operation_Status -OperationStatus $script:operationStatus -Key 'IPConfiguration' -Status 'Successful' -Message 'DHCP configuration applied.' -Details "Adapter: $adapterName"
		}
		else
		{
			Write_MgLog ("Failed to configure DHCP: {0}" -f (($out | Out-String).Trim())) 'fail'
			Set_Operation_Status -OperationStatus $script:operationStatus -Key 'IPConfiguration' -Status 'Failed' -Message 'Failed to configure DHCP.' -Details (($out | Out-String).Trim())
		}
		return
	}
	$ipAddress = $Net.Ip
	$gateway = $Net.Gateway
	$out = netsh interface ip set address name="$adapterName" static $ipAddress 255.255.255.0 $gateway 2>&1
	if ($LASTEXITCODE -ne 0)
	{
		Write_MgLog ("Failed to configure static IP: {0}" -f (($out | Out-String).Trim())) 'fail'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'IPConfiguration' -Status 'Failed' -Message 'Failed to configure static IP.' -Details (($out | Out-String).Trim())
		return
	}
	try { Set-DnsClientServerAddress -InterfaceAlias $adapterName -ServerAddresses ("8.8.8.8", "8.8.4.4") -ErrorAction Stop }
	catch { Write_MgLog "IP set, but DNS could not be set: $($_.Exception.Message)" 'warn' }
	Write_MgLog "Static IP $ipAddress / 255.255.255.0, gateway $gateway, DNS 8.8.8.8 + 8.8.4.4 on '$adapterName'." 'ok'
	Set_Operation_Status -OperationStatus $script:operationStatus -Key 'IPConfiguration' -Status 'Successful' -Message 'Static IP configuration applied.' -Details "Adapter: $adapterName, IP: $ipAddress"
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Invoke_MiniGhost_Full  (runspace)
#   Same calls as the Mini Ghost buttons, in dependency order:
#   1 Update Store Number   2 Change Machine Name (+ old X/Z folders, INI + Startup.ini + SQL sync)
#   3 Truncate Tables       4 Registry Cleanup     5 Configure Network
# ---------------------------------------------------------------------------------------------------
function Invoke_MiniGhost_Full
{
	param ([string]$NewStore, [string]$TargetName, $Net, [string[]]$Tables)
	$oldName = $env:COMPUTERNAME
	$failures = 0
	$renamed = $false

	$oldStore = Get_Store_Number_From_INI
	if (-not $oldStore)
	{
		Write_MgLog 'Store number not found in Startup.ini - nothing was changed.' 'fail'
		return @{ Renamed = $false; Failures = 1 }
	}

	# ---- 1. store number (Mini Ghost "Update Store Number")
	Write_MgLog ('[1/5] Store number: {0} -> {1}' -f $oldStore, $NewStore) 'step'
	if ($oldStore -eq $NewStore)
	{
		Write_MgLog "Store number is already $NewStore - the INIs are refreshed by the machine sync (step 2)." 'muted'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'StoreNumberChange' -Status 'Skipped' -Message 'Store number unchanged.'
	}
	else
	{
		$iniOk = $false
		try { $iniOk = [bool](Update_INIs -newStoreNumber $NewStore -MachineName $TargetName -OldStoreNumber $oldStore -OldMachineName $oldName -CreateSmsHttpsIfMissing) }
		catch { Write_MgLog "Update_INIs error: $($_.Exception.Message)" 'fail' }
		if ($iniOk)
		{
			$script:newStoreNumber = $NewStore
			Write_MgLog "Store number changed to $NewStore in the INIs." 'ok'
			Set_Operation_Status -OperationStatus $script:operationStatus -Key 'StoreNumberChange' -Status 'Successful' -Message 'Store number updated in INIs.' -Details "Store number changed to '$NewStore'."
		}
		else
		{
			$failures++
			Write_MgLog 'Update_INIs failed (best-effort) - see the lines above.' 'fail'
			Set_Operation_Status -OperationStatus $script:operationStatus -Key 'StoreNumberChange' -Status 'Failed' -Message 'Update_INIs failed.' -Details 'Best-effort update encountered an error.'
		}
		$sqlStore = Update_SQL_Tables_For_Store_Number_Change -storeNumber $NewStore
		if ($sqlStore.Success) { Write_MgLog 'STD_TAB / TER_TAB updated with the new store number.' 'ok' }
		else
		{
			$failures++
			Write_MgLog ('SQL store number update failed: {0}' -f (@($sqlStore.FailedCommands) -join ' | ')) 'fail'
		}
	}

	# ---- 2. machine name (Mini Ghost "Change Machine Name" workflow, without the dialogs)
	Write_MgLog ('[2/5] Machine name: {0} -> {1}' -f $oldName, $TargetName) 'step'
	$ctx = Get_Machine_Name_Context -MachineName $TargetName -CurrentMachineName $oldName
	if (-not $ctx.Success)
	{
		$failures++
		Write_MgLog ("Invalid machine name: {0}" -f $ctx.ErrorMessage) 'fail'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'MachineNameChange' -Status 'Failed' -Message 'Invalid machine name.' -Details $ctx.ErrorMessage
	}
	else
	{
		$store = Get_Store_Number_From_INI
		$startup = Resolve_Startup_Ini_Path -StartupIniPath $StartupIniPath
		$script:newMachineName = $ctx.MachineName
		$doSync = $true
		if (-not $ctx.IsSameMachineName)
		{
			try
			{
				Rename-Computer -NewName $ctx.MachineName -Force -ErrorAction Stop
				$renamed = $true
				Write_MgLog "Computer renamed to $($ctx.MachineName) (active after the reboot)." 'ok'
				Set_Operation_Status -OperationStatus $script:operationStatus -Key 'MachineNameChange' -Status 'Successful' -Message 'Machine name changed successfully.' -Details "Machine name changed to '$($ctx.MachineName)'."
				$null = Remove_Old_XZ_Folders -MachineName $ctx.MachineName -StoreNumber $store -OperationStatus $script:operationStatus
			}
			catch
			{
				$doSync = $false
				$failures++
				Write_MgLog "Error changing machine name: $($_.Exception.Message)" 'fail'
				Set_Operation_Status -OperationStatus $script:operationStatus -Key 'MachineNameChange' -Status 'Failed' -Message 'Error changing machine name.' -Details "Error: $($_.Exception.Message)"
			}
		}
		else
		{
			Write_MgLog "Machine name is already $($ctx.MachineName) - running the full INI + SQL sync without renaming." 'info'
		}

		if ($doSync)
		{
			$sync = Invoke_Machine_Name_Sync -StoreNumber $store -MachineName $ctx.MachineName -OldMachineName $oldName -MachineContext $ctx -StartupIniPath $startup -OperationStatus $script:operationStatus
			if ($sync.IniSyncSuccess) { Write_MgLog 'INI files synced.' 'ok' } else { $failures++; Write_MgLog "INI sync failed: $($sync.IniSyncError)" 'fail' }
			if ($sync.StartupIniSuccess) { Write_MgLog "Startup.ini: $($sync.StartupIniDetails)" 'ok' } else { $failures++; Write_MgLog "Startup.ini update failed: $($sync.StartupIniError)" 'fail' }
			if ($sync.SqlSyncSuccess) { Write_MgLog 'SQL tables synced (TER/STO/LNK/RUN/STD).' 'ok' }
			else
			{
				$failures++
				Write_MgLog "SQL sync failed: $($sync.SqlSyncError)" 'fail'
				if ($sync.SqlResult -and $sync.SqlResult.FailedCommands) { foreach ($fc in @($sync.SqlResult.FailedCommands)) { Write_MgLog $fc 'muted' } }
			}
			if ($ctx.IsSameMachineName)
			{
				$st = if ($sync.Success) { 'Successful' } else { 'Failed' }
				Set_Operation_Status -OperationStatus $script:operationStatus -Key 'MachineNameChange' -Status $st -Message 'Same-name INI + SQL sync.' -Details "Machine name already '$($ctx.MachineName)'."
			}
		}
	}

	# ---- 3. truncate (always)
	Write_MgLog ('[3/5] Truncate {0} tables' -f $Tables.Count) 'step'
	$failedTables = @(RPT_Blank -tables $Tables)
	if ($failedTables.Count -eq 0)
	{
		Write_MgLog 'Tables truncated successfully.' 'ok'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'TableTruncation' -Status 'Successful' -Message 'Tables truncated successfully.'
	}
	else
	{
		$failures++
		Write_MgLog ('Failed to truncate {0} table(s): {1}' -f $failedTables.Count, ($failedTables -join ', ')) 'fail'
		if ($script:LastSqlExecutionError) { Write_MgLog "Last SQL error: $script:LastSqlExecutionError" 'muted' }
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'TableTruncation' -Status 'Failed' -Message 'Failed to truncate some tables.' -Details "Failed tables: $($failedTables -join ', ')"
	}

	# ---- 4. registry GT* counters
	Write_MgLog '[4/5] Registry cleanup (GT* counters)' 'step'
	$gt = Remove_GT_Registry_Values
	if ($gt.Status -eq 'Successful')
	{
		Write_MgLog ("{0} 'GT' registry value(s) deleted." -f $gt.DeletedCount) 'ok'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'RegistryCleanup' -Status 'Successful' -Message 'GT registry values removed successfully.' -Details "$($gt.DeletedCount) 'GT' registry keys were deleted."
	}
	else
	{
		$failures++
		Write_MgLog "Registry cleanup failed: $($gt.Message)" 'fail'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'RegistryCleanup' -Status 'Failed' -Message 'Failed to remove GT registry values.' -Details $gt.Message
	}

	# ---- 5. network (last: on site, the connection may drop)
	Write_MgLog '[5/5] Network' 'step'
	Set_MiniGhost_Network $Net
	if ($script:operationStatus['IPConfiguration'].Status -eq 'Failed') { $failures++ }

	Write_MgSummary
	return @{ Renamed = $renamed; NewName = $TargetName; Failures = $failures }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Invoke_MiniGhost_DbUpdate  (runspace)
#   Database copied from another lane: only the SQL sync for THIS lane (Mini Ghost "Update SQL
#   Database") + the truncate. -SourceTerminal = lane the database came from (RUN_TAB swap).
# ---------------------------------------------------------------------------------------------------
function Invoke_MiniGhost_DbUpdate
{
	param ([string]$SourceTerminal, [string[]]$Tables)
	$failures = 0
	$store = Get_Store_Number_From_INI
	if (-not $store)
	{
		Write_MgLog 'Store number not found in Startup.ini - nothing was changed.' 'fail'
		return @{ Failures = 1 }
	}
	$ctx = Get_Machine_Name_Context -MachineName $env:COMPUTERNAME -CurrentMachineName $env:COMPUTERNAME
	if (-not $ctx.Success)
	{
		Write_MgLog "This machine's name could not be read as a lane name: $($ctx.ErrorMessage)" 'fail'
		return @{ Failures = 1 }
	}

	Write_MgLog ('[1/2] SQL sync for store {0}, lane {1} ({2}); database copied from lane {3}' -f $store, $ctx.MachineNumber, $ctx.MachineName, $SourceTerminal) 'step'
	$r = Update_SQL_Tables_For_Machine_Name_Change -storeNumber $store -machineName $ctx.MachineName -machineNumber $ctx.MachineNumber -OldMachineName $env:COMPUTERNAME -OldMachineNumber $SourceTerminal
	if ($r -and $r.Success)
	{
		$details = 'Protected SQL sync completed.'
		if ($r.ContainsKey('RunTabUpdated')) { $details += " RUN updated: $($r.RunTabUpdated)." }
		if ($r.ContainsKey('LocalTerminalDB')) { $details += " Local terminal DB: $($r.LocalTerminalDB)." }
		Write_MgLog "SQL database updated successfully. $details" 'ok'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'SQLDatabaseUpdate' -Status 'Successful' -Message 'SQL database updated successfully.' -Details $details
	}
	else
	{
		$failures++
		Write_MgLog 'Failed to update the SQL database.' 'fail'
		if ($r -and $r.FailedCommands) { foreach ($fc in @($r.FailedCommands)) { Write_MgLog $fc 'muted' } }
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'SQLDatabaseUpdate' -Status 'Failed' -Message 'Failed to update SQL database.'
	}

	Write_MgLog ('[2/2] Truncate {0} tables' -f $Tables.Count) 'step'
	$failedTables = @(RPT_Blank -tables $Tables)
	if ($failedTables.Count -eq 0)
	{
		Write_MgLog 'Tables truncated successfully.' 'ok'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'TableTruncation' -Status 'Successful' -Message 'Tables truncated successfully.'
	}
	else
	{
		$failures++
		Write_MgLog ('Failed to truncate {0} table(s): {1}' -f $failedTables.Count, ($failedTables -join ', ')) 'fail'
		Set_Operation_Status -OperationStatus $script:operationStatus -Key 'TableTruncation' -Status 'Failed' -Message 'Failed to truncate some tables.' -Details "Failed tables: $($failedTables -join ', ')"
	}
	Write_MgSummary
	return @{ Failures = $failures }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Start_MgJob - runs a workflow + the core functions in a background runspace
# ---------------------------------------------------------------------------------------------------
function Get_MgRunspaceScript
{
	$defs = foreach ($n in $script:MgRunspaceFunctions) { "function $n {`r`n" + (Get-Item "function:$n").ScriptBlock.ToString() + "`r`n}" }
	"param(`$Root, `$Workflow, `$Params)`r`n" + ($defs -join "`r`n") + "`r`nInitialize_MiniGhost_Context -StoremanRoot `$Root`r`n& `$Workflow @Params"
}

function Start_MgJob
{
	param ([string]$Workflow, [hashtable]$Params, [string]$OnDone)
	$ps = [PowerShell]::Create()
	$null = $ps.AddScript((Get_MgRunspaceScript)).AddArgument($script:Db.Storeman).AddArgument($Workflow).AddArgument($Params)
	$script:MgJob = @{ PS = $ps; Handle = $ps.BeginInvoke(); Started = [datetime]::Now; Info = 0; Warn = 0; Err = 0; OnDone = $OnDone }
	Set_RunningUi $true
	$script:UI.MgPump.Start()
}

# Timer tick: Write-Host / warnings / errors of the runspace -> terminal
function Update_MgJobOutput
{
	$j = $script:MgJob
	if (-not $j) { $script:UI.MgPump.Stop(); return }
	$finished = $j.Handle.IsCompleted
	$s = $j.PS.Streams
	while ($j.Info -lt $s.Information.Count)
	{
		$md = $s.Information[$j.Info].MessageData
		$j.Info++
		$text = [string]$md
		$kind = 'text'
		if ($md -is [System.Management.Automation.HostInformationMessage])
		{
			$text = [string]$md.Message
			$kind = switch ([string]$md.ForegroundColor) { 'Green' { 'ok' } 'Red' { 'fail' } 'Yellow' { 'warn' } 'DarkYellow' { 'warn' } 'Cyan' { 'info' } 'Magenta' { 'step' } 'DarkGray' { 'muted' } 'Gray' { 'muted' } default { 'text' } }
		}
		if ($text.Trim()) { Write_Term $text.TrimEnd() $kind }
	}
	while ($j.Warn -lt $s.Warning.Count) { Write_Term ([string]$s.Warning[$j.Warn].Message) 'warn'; $j.Warn++ }
	while ($j.Err -lt $s.Error.Count) { Write_Term ([string]$s.Error[$j.Err].Exception.Message) 'fail'; $j.Err++ }
	$script:UI.Elapsed.Text = '{0:hh\:mm\:ss}' -f ([datetime]::Now - $j.Started)

	if ($finished)
	{
		$script:UI.MgPump.Stop()
		$result = $null
		try { $result = @($j.PS.EndInvoke($j.Handle)) | Where-Object { $_ -is [hashtable] -and $_.ContainsKey('Failures') } | Select-Object -Last 1 }
		catch { Write_Term ('Mini Ghost error: {0}' -f $_.Exception.Message) 'fail' }
		$j.PS.Dispose()
		$script:MgJob = $null
		Set_RunningUi $false
		Set_Status 'Finished' 100
		if ($j.OnDone) { & $j.OnDone $result }
	}
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Read_Prompt - asks one question in the prompt bar under the terminal (Enter = OK,
#   Esc = Cancel). Returns the answer, or $null when cancelled. -Pattern re-asks on bad input.
# ---------------------------------------------------------------------------------------------------
function Read_Prompt
{
	param ([string]$Question, [string]$Default = '', [string]$Pattern, [string]$Hint = '')
	$ui = $script:UI
	$ui.PromptLabel.Text = $Question
	$ui.PromptHint.Text = $Hint
	$ui.PromptBox.Text = $Default
	$ui.PromptBar.Visible = $true
	$null = $ui.PromptBox.Focus()   # Focus() returns a bool: it must not become part of the answer
	$ui.PromptBox.SelectAll()
	try
	{
		while ($true)
		{
			$ui.PromptAnswer = $null
			$ui.PromptCancel = $false
			while ($null -eq $ui.PromptAnswer -and -not $ui.PromptCancel)
			{
				if ($ui.Form.IsDisposed -or $ui.Closing) { return $null }
				[System.Windows.Forms.Application]::DoEvents()
				Start-Sleep -Milliseconds 15
			}
			if ($ui.PromptCancel) { Write_Term ('{0}  -> cancelled' -f $Question) 'warn'; return $null }
			$a = ([string]$ui.PromptAnswer).Trim()
			if (-not $Pattern -or $a -match $Pattern)
			{
				Write_Term ('{0}  {1}' -f $Question, $(if ($a) { $a } else { '(Enter)' })) 'muted'
				return $a
			}
			Write_Term ("Invalid answer '{0}'. {1}" -f $a, $Hint) 'warn'
			$ui.PromptBox.SelectAll()
		}
	}
	finally { if (-not $ui.Form.IsDisposed) { $ui.PromptBar.Visible = $false } }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Show_MiniGhost_Info - lane details in the terminal (Store #, Lane #, name, IP)
# ---------------------------------------------------------------------------------------------------
function Show_MiniGhost_Info
{
	Write_Term 'Mini Ghost - lane details' 'step'
	Set_Status 'Reading lane details...'
	[System.Windows.Forms.Application]::DoEvents()
	Initialize_MiniGhost_Context -StoremanRoot $script:Db.Storeman
	$store = Get_Store_Number_From_INI
	Get_Store_Name_From_INI
	$ter = $null
	try { $line = Get-Content -LiteralPath $script:StartupIniPath -ErrorAction Stop | Where-Object { $_ -match '^\s*TER\s*=\s*\d+' } | Select-Object -First 1; if ($line -match '=\s*(\d+)') { $ter = ([int]$Matches[1]).ToString('000') } }
	catch { }
	$adapters = @(Get_Active_IP_Config | Where-Object { $_ })
	$script:MgInfo = @{ Store = $store; Terminal = $ter; Machine = $env:COMPUTERNAME; Adapters = $adapters }

	Write_Term ('Machine name : {0}' -f $env:COMPUTERNAME) 'text' -NoTime
	Write_Term ('Store #      : {0}   {1}' -f $(if ($store) { $store } else { 'not found in Startup.ini' }), $(if ($script:FunctionResults['StoreName'] -and $script:FunctionResults['StoreName'] -ne 'N/A') { '(' + $script:FunctionResults['StoreName'] + ')' } else { '' })) 'text' -NoTime
	Write_Term ('Lane #       : {0}' -f $(if ($ter) { $ter } else { 'TER= not found in Startup.ini' })) 'text' -NoTime
	if ($adapters.Count -eq 0) { Write_Term 'IP address   : no active network adapter with an IPv4 address' 'warn' -NoTime }
	foreach ($a in $adapters)
	{
		Write_Term ('IP address   : {0}   gateway {1}   ({2})' -f @($a.IPv4Address)[0].IPAddress, $a.IPv4DefaultGateway.NextHop, $a.InterfaceAlias) 'text' -NoTime
	}
	if ($script:Db -and $script:Db.Ok) { Write_Term ('Database     : {0} on {1}' -f $script:Db.Database, $script:Db.Server) 'text' -NoTime }
	Set_Status 'Mini Ghost'
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Start_MiniGhost_Run - prompts (store #, lane #, IP), preview, confirm, run
# ---------------------------------------------------------------------------------------------------
function Start_MiniGhost_Run
{
	if ($script:MgJob -or $script:UI.Prompting) { return }
	$script:UI.Prompting = $true
	Set_RunningUi $true
	try
	{
		if (-not $script:MgInfo) { Show_MiniGhost_Info }
		$info = $script:MgInfo
		Write_Term 'Run Mini-Ghost' 'step'

		# ---- store #
		$store = Read_Prompt 'Store #:' ([string]$info.Store) '^\d{3,4}$' '3 or 4 digits (Enter keeps the current store).'
		if ($null -eq $store) { return }

		# ---- lane # -> machine name (current naming style, like typing 001 in Mini Ghost)
		$ctx = $null
		while (-not $ctx)
		{
			$lane = Read_Prompt 'Lane #:' ([string]$info.Terminal) '^\d{1,3}$' 'Lane number, 1-3 digits (e.g. 5 or 005).'
			if ($null -eq $lane) { return }
			$lane = ([int]$lane).ToString('000')
			if ($lane -eq '000' -or [int]$lane -ge 900) { Write_Term 'Lane # must be 001-899.' 'warn'; continue }
			$c = Get_Machine_Name_Context -MachineName $lane -CurrentMachineName $env:COMPUTERNAME
			if (-not $c.Success)
			{
				Write_Term ("The current name '{0}' has no reusable lane style: {1}" -f $env:COMPUTERNAME, $c.ErrorMessage) 'warn'
				$full = Read_Prompt ('Full machine name for lane {0}:' -f $lane) '' '^[A-Za-z0-9-]{1,15}$' 'e.g. POS005, LN026005, 0384LANE005 (max 15 characters).'
				if ($null -eq $full) { return }
				$c = Get_Machine_Name_Context -MachineName $full -CurrentMachineName $env:COMPUTERNAME
				if (-not $c.Success) { Write_Term $c.ErrorMessage 'warn'; continue }
			}
			$ctx = $c
		}

		# ---- IP
		$net = @{ Mode = 'Keep' }
		$adapters = @($info.Adapters)
		$adapter = $null
		if ($adapters.Count -gt 1)
		{
			for ($i = 0; $i -lt $adapters.Count; $i++) { Write_Term ('  {0}) {1} - IP {2}' -f ($i + 1), $adapters[$i].InterfaceAlias, @($adapters[$i].IPv4Address)[0].IPAddress) 'text' -NoTime }
			$pick = Read_Prompt 'Network adapter #:' '1' ('^[1-{0}]$' -f [Math]::Min(9, $adapters.Count)) 'Number from the list above.'
			if ($null -eq $pick) { return }
			$adapter = $adapters[[int]$pick - 1]
		}
		elseif ($adapters.Count -eq 1) { $adapter = $adapters[0] }

		if ($adapter)
		{
			$curIp = [string]@($adapter.IPv4Address)[0].IPAddress
			$gw = [string]$adapter.IPv4DefaultGateway.NextHop
			$prefix = if ($gw -match '^(\d{1,3}\.\d{1,3}\.\d{1,3})\.\d{1,3}$') { $Matches[1] } else { $null }
			while ($true)
			{
				$ans = Read_Prompt ('IP (current {0}):' -f $curIp) '' '^(|dhcp|DHCP|\d{1,3}|\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})$' "Enter = keep $curIp, 'dhcp', last octet (1-254) or full IP."
				if ($null -eq $ans) { return }
				if ($ans -eq '') { $net = @{ Mode = 'Keep' }; break }
				if ($ans -ieq 'dhcp') { $net = @{ Mode = 'Dhcp'; Adapter = $adapter.InterfaceAlias }; break }
				if (-not $prefix) { Write_Term "No usable gateway ($gw) on '$($adapter.InterfaceAlias)' - only Enter or dhcp are possible." 'warn'; continue }
				$ip = if ($ans -match '^\d{1,3}$') { "$prefix.$ans" } else { $ans }
				$last = [int]($ip.Split('.')[3])
				if ($last -lt 1 -or $last -gt 254 -or @($ip.Split('.') | Where-Object { [int]$_ -gt 255 }).Count) { Write_Term "Invalid IP '$ip'." 'warn'; continue }
				if ($ip -eq $curIp) { $net = @{ Mode = 'Keep' }; break }
				if (-not $ip.StartsWith("$prefix.")) { Write_Term "Note: $ip is not in the gateway's network ($prefix.x) - mask stays 255.255.255.0." 'warn' }
				Set_Status "Checking that $ip is free..."
				[System.Windows.Forms.Application]::DoEvents()
				if (Test-Connection -ComputerName $ip -Count 1 -Quiet) { Write_Term "The IP address '$ip' is already in use. Please choose a different one." 'warn'; continue }
				$net = @{ Mode = 'Static'; Adapter = $adapter.InterfaceAlias; Ip = $ip; Gateway = $gw }
				break
			}
		}
		else { Write_Term 'No active network adapter - the IP is left unchanged.' 'warn' }

		# ---- preview + confirm
		$ipText = switch ($net.Mode) { 'Keep' { 'unchanged' } 'Dhcp' { "DHCP on $($net.Adapter)" } default { "$($net.Ip) /24, gateway $($net.Gateway), DNS 8.8.8.8 + 8.8.4.4 ($($net.Adapter))" } }
		Write_Term 'Mini Ghost will:' 'step'
		Write_Term ('  Store #      {0} -> {1}' -f $info.Store, $store) 'text' -NoTime
		Write_Term ('  Lane #       {0} -> {1}' -f $info.Terminal, $ctx.MachineNumber) 'text' -NoTime
		Write_Term ('  Machine name {0} -> {1}{2}' -f $env:COMPUTERNAME, $ctx.MachineName, $(if ($ctx.IsSameMachineName) { '  (same name: sync only)' } else { '' })) 'text' -NoTime
		Write_Term ('  IP           {0}' -f $ipText) 'text' -NoTime
		Write_Term ('  + INI files, Startup.ini, SQL (TER/STO/LNK/RUN/STD), old X/Z folders, truncate {0} tables, GT registry counters' -f $script:MgTruncateTables.Count) 'text' -NoTime
		$ok = Read_Prompt 'Run Mini Ghost now? (Y/N)' 'Y' '^(?i)(y|yes|n|no)$' 'Y or N.'
		if ($ok -notmatch '^(?i)y') { Write_Term 'Mini Ghost cancelled - nothing was changed.' 'warn'; return }

		Write_Term ('==== Mini Ghost: {0} / store {1} / lane {2} ====' -f $ctx.MachineName, $store, $ctx.MachineNumber) 'info'
		$script:UI.Prompting = $false
		Start_MgJob 'Invoke_MiniGhost_Full' @{ NewStore = $store; TargetName = $ctx.MachineName; Net = $net; Tables = $script:MgTruncateTables } 'Complete_MiniGhost_Run'
	}
	finally
	{
		if ($script:UI.Prompting) { $script:UI.Prompting = $false; Set_RunningUi $false }
	}
}

# After the run: ask to reboot, like Mini Ghost
function Complete_MiniGhost_Run
{
	param ($Result)
	$fails = if ($Result) { [int]$Result.Failures } else { 1 }
	if ($fails -eq 0) { Write_Term 'Mini Ghost finished without errors.' 'ok' }
	else { Write_Term ("Mini Ghost finished with {0} error(s) - see above. A reboot is still needed for the new name." -f $fails) 'warn' }
	$script:UI.Prompting = $true
	Set_RunningUi $true
	try
	{
		$ans = Read_Prompt 'Reboot now? (Y/N)' 'Y' '^(?i)(y|yes|n|no)$' 'Y or N.'
		if ($ans -match '^(?i)y')
		{
			Write_Term 'Rebooting...' 'info'
			[System.Windows.Forms.Application]::DoEvents()
			Restart-Computer -Force
		}
		else { Write_Term 'Reboot postponed - restart the lane before using it.' 'warn' }
	}
	finally { $script:UI.Prompting = $false; Set_RunningUi $false }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Get_MiniGhost_SourceTerminal - lane # the copied database belongs to (from RUN_TAB)
# ---------------------------------------------------------------------------------------------------
function Get_MiniGhost_SourceTerminal
{
	$found = @()
	$cn = New-Object System.Data.SqlClient.SqlConnection $script:Db.ConnectionString
	try
	{
		$cn.Open()
		$cmd = $cn.CreateCommand()
		$cmd.CommandText = "IF OBJECT_ID('RUN_TAB','U') IS NOT NULL SELECT DISTINCT LTRIM(RTRIM(CONVERT(varchar(10), F1000))) FROM RUN_TAB"
		$rd = $cmd.ExecuteReader()
		while ($rd.Read()) { $v = [string]$rd.GetValue(0); if ($v -match '^\d{1,3}$' -and [int]$v -gt 0 -and [int]$v -lt 900) { $found += ([int]$v).ToString('000') } }
		$rd.Close()
	}
	catch { }
	finally { $cn.Dispose() }
	@($found | Select-Object -Unique)
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Start_MiniGhost_DbUpdate - "Update Store SQL Database" (copied database)
# ---------------------------------------------------------------------------------------------------
function Start_MiniGhost_DbUpdate
{
	if ($script:MgJob -or $script:UI.Prompting) { return }
	if (-not ($script:Db -and $script:Db.Ok)) { Write_Term 'Not connected to the lane database.' 'fail'; return }
	$script:UI.Prompting = $true
	Set_RunningUi $true
	try
	{
		if (-not $script:MgInfo) { Show_MiniGhost_Info }
		$info = $script:MgInfo
		Write_Term 'Update Store SQL Database' 'step'
		$ctx = Get_Machine_Name_Context -MachineName $env:COMPUTERNAME -CurrentMachineName $env:COMPUTERNAME
		if (-not $ctx.Success) { Write_Term ("This machine's name '{0}' is not a lane name: {1}" -f $env:COMPUTERNAME, $ctx.ErrorMessage) 'fail'; return }

		$sources = @(Get_MiniGhost_SourceTerminal | Where-Object { $_ -ne $ctx.MachineNumber })
		$guess = if ($sources.Count -eq 1) { $sources[0] } else { '' }
		if ($sources.Count) { Write_Term ('Terminal(s) found in the copied database (RUN_TAB): {0}' -f ($sources -join ', ')) 'muted' }
		$src = Read_Prompt 'Lane # the database was copied FROM:' $guess '^\d{1,3}$' 'The source lane number (e.g. 3 or 003).'
		if ($null -eq $src) { return }
		$src = ([int]$src).ToString('000')

		Write_Term 'This will update the lane database only:' 'step'
		Write_Term ('  Store {0}, lane {1} ({2}) - database from lane {3}' -f $info.Store, $ctx.MachineNumber, $ctx.MachineName, $src) 'text' -NoTime
		Write_Term ('  SQL sync of TER/STO/LNK/RUN/STD + truncate {0} tables. No rename, no INI, no IP changes.' -f $script:MgTruncateTables.Count) 'text' -NoTime
		$ok = Read_Prompt 'Update the database now? (Y/N)' 'Y' '^(?i)(y|yes|n|no)$' 'Y or N.'
		if ($ok -notmatch '^(?i)y') { Write_Term 'Cancelled - nothing was changed.' 'warn'; return }

		$script:UI.Prompting = $false
		Start_MgJob 'Invoke_MiniGhost_DbUpdate' @{ SourceTerminal = $src; Tables = $script:MgTruncateTables } 'Complete_MiniGhost_DbUpdate'
	}
	finally
	{
		if ($script:UI.Prompting) { $script:UI.Prompting = $false; Set_RunningUi $false }
	}
}

function Complete_MiniGhost_DbUpdate
{
	param ($Result)
	if ($Result -and [int]$Result.Failures -eq 0) { Write_Term 'Lane database updated. Restart SMS on this lane.' 'ok' }
	else { Write_Term 'Database update finished with errors - see above.' 'warn' }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Update_MiniGhostNav - Mini Ghost is a lane tool: greyed out unless this is a lane
# ---------------------------------------------------------------------------------------------------
function Update_MiniGhostNav
{
	# Mini Ghost = lanes only; operation tools (Scales Maintenance, Remote Lane Tools) = their Role
	$role = if ($script:Db) { [string]$script:Db.Role } else { '' }
	$lane = ($role -eq 'Lane')
	$script:UI.NavMg.Enabled = $lane
	$script:UI.NavMg.ForeColor = if ($lane) { $script:T.Text } else { $script:T.Muted }
	$leave = (-not $lane -and $script:Tool -eq 'MiniGhost')
	foreach ($k in $script:OpsTools.Keys)
	{
		# Role: Server / Lane = that machine only; Known = lane or server; Any = always
		$need = $script:OpsTools[$k].Role
		$on = ($need -eq 'Any') -or ($need -eq 'Known' -and $role -in @('Lane', 'Server')) -or ($role -eq $need)
		$script:UI.Nav[$k].Button.Enabled = $on
		$script:UI.Nav[$k].Button.ForeColor = if ($on) { $script:T.Text } else { $script:T.Muted }
		if (-not $on -and $script:Tool -eq $k) { $leave = $true }
	}
	$tip = if ($lane) { 'Lane: Mini Ghost is available. Scales Maintenance, Remote Lane Tools and Export are for the store server.' }
	elseif ($role -eq 'Server') { 'Store server: Scales Maintenance, Remote Lane Tools and Export are available. Mini Ghost is for lanes only.' }
	else { 'The lane / server tools need a detected machine role. Repair Windows always works.' }
	if ($script:UI.SmsFn) { $script:UI.SmsFn.Visible = ($role -eq 'Server') }
	$script:UI.Tip.SetToolTip($script:UI.NavMg.Parent, $tip)   # disabled buttons show no tooltip
	if ($leave) { Select_Tool 'Db' }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Select_Tool - switches the content area between the tools
# ---------------------------------------------------------------------------------------------------
function Select_Tool
{
	param ([string]$Tool)
	if ($script:Run -and -not $script:Run.Sync.Done) { return }
	if ($script:MgJob -or $script:UI.Prompting) { return }
	if (-not $script:UI.Nav[$Tool] -or -not $script:UI.Nav[$Tool].Button.Enabled) { return }
	if ($script:OpsTools.Contains([string]$script:Tool) -and $Tool -ne $script:Tool) { Save_OpsChecked $script:Tool }
	$script:Tool = $Tool
	$isDb = ($Tool -eq 'Db'); $isMg = ($Tool -eq 'MiniGhost')
	$script:UI.Toolbar.Visible = $isDb
	$script:UI.ListCard.Visible = $isDb
	$script:UI.ListGap.Visible = $isDb
	$script:UI.MgPanel.Visible = $isMg
	foreach ($k in $script:UI.Ops.Keys)
	{
		$o = $script:UI.Ops[$k]
		$o.Toolbar.Visible = ($k -eq $Tool); $o.Card.Visible = ($k -eq $Tool); $o.Gap.Visible = ($k -eq $Tool)
	}
	foreach ($k in $script:UI.Nav.Keys)
	{
		$script:UI.Nav[$k].Button.BackColor = if ($k -eq $Tool) { $script:T.CardHi } else { $script:T.Side }
		$script:UI.Nav[$k].Mark.Visible = ($k -eq $Tool)
	}
	$script:UI.Sub.Text = switch ($Tool) { 'Db' { '|  Database Maintenance' } 'MiniGhost' { '|  Mini Ghost' } default { '|  ' + $script:OpsTools[$Tool].Title } }
	if ($isMg) { Show_MiniGhost_Info }
	elseif (-not $isDb) { & $script:OpsTools[$Tool].InfoFn }
}

# ===================================================================================================
#                  SECTION: Scales core (ported verbatim from the store maintenance script)
# ---------------------------------------------------------------------------------------------------
#   Source: Desktop\script.ps1 (Alex_C.T). Business logic UNCHANGED. They log through Write_Log (the
#   toolkit shim in "Scales Maintenance tool" sends it to the terminal) and read the $script: paths
#   set by Initialize_ScalesContext.
# ===================================================================================================

# ===================================================================================================
#                              FUNCTION: Normalize_Sql_ConnectionString
# ---------------------------------------------------------------------------------------------------
# Description:
#   Turns a stored connection string (or the Server/Database/User/Pwd parts) into one clean
#   SqlClient connection string. Treats 'N/A' placeholders as empty and always adds
#   TrustServerCertificate=True so older Windows builds can reach newer SQL servers.
#
# Parameters:
#   - RawConnectionString  [any]: value from FunctionResults (string, hashtable or object with .ConnectionString)
#   - Server               [string]: SQL server name (used when no usable connection string is stored)
#   - Database             [string]: database name (used when no usable connection string is stored)
#   - SqlUser              [string]: SQL login (used with Server/Database)
#   - SqlPwd               [string]: SQL login password (used with Server/Database)
#
# Returns:
#   The normalized connection string, or $null when nothing usable was given.
# ===================================================================================================

function Normalize_Sql_ConnectionString
{
	param (
		[Parameter(Mandatory = $false)]
		$RawConnectionString,
		[string]$Server,
		[string]$Database,
		[string]$SqlUser,
		[string]$SqlPwd
	)

	$connStr = $null
	try
	{
		if ($RawConnectionString -is [hashtable] -and $RawConnectionString.ContainsKey('ConnectionString'))
		{
			$connStr = [string]$RawConnectionString['ConnectionString']
		}
		elseif ($null -ne $RawConnectionString -and $RawConnectionString.PSObject -and $RawConnectionString.PSObject.Properties['ConnectionString'])
		{
			$connStr = [string]$RawConnectionString.ConnectionString
		}
		else
		{
			$connStr = [string]$RawConnectionString
		}
	}
	catch
	{
		try { $connStr = [string]$RawConnectionString } catch { $connStr = $null }
	}

	if (-not [string]::IsNullOrWhiteSpace($connStr))
	{
		$connStr = $connStr.Trim().Trim('"')
		if ($connStr -match '^\s*ConnectionString\s*=\s*(.+?)\s*$')
		{
			$connStr = $matches[1].Trim()
		}
	}

	if ([string]::IsNullOrWhiteSpace($connStr) -or $connStr -eq 'N/A' -or $connStr -eq 'System.Collections.Hashtable')
	{
		$connStr = $null
	}

	# Treat the 'N/A' placeholder the same as empty - otherwise this builds a
	# real connection string pointing at a SQL server literally named "N/A",
	# and every caller pays a full DNS/connect timeout (~15s at startup).
	$serverNorm = $Server
	$databaseNorm = $Database
	try
	{
		if ([string]$Server -eq 'N/A' -or [string]$Server -eq 'System.Collections.Hashtable') { $serverNorm = '' }
		if ([string]$Database -eq 'N/A' -or [string]$Database -eq 'System.Collections.Hashtable') { $databaseNorm = '' }
	}
	catch { }

	if (-not $connStr -and -not [string]::IsNullOrWhiteSpace($serverNorm) -and -not [string]::IsNullOrWhiteSpace($databaseNorm))
	{
		if (-not [string]::IsNullOrWhiteSpace($SqlUser) -and -not [string]::IsNullOrWhiteSpace($SqlPwd))
		{
			$connStr = "Server=${serverNorm};Database=${databaseNorm};User ID=$SqlUser;Password=$SqlPwd;TrustServerCertificate=True;"
		}
		else
		{
			$connStr = "Server=${serverNorm};Database=${databaseNorm};Integrated Security=SSPI;TrustServerCertificate=True;"
		}
	}

	# Append TrustServerCertificate=True when it is not set. Plain text on purpose: a
	# SqlConnectionStringBuilder would rewrite Server=/Database= to Data Source=/Initial Catalog=
	# (callers parse those), and PowerShell treats the builder as a dictionary, so
	# '$csb.ConnectionString = ...' throws instead of loading (setting a builder property needs
	# '$csb.PSBase.<Property> = ...') - before v2.7.0 that made this step a silent no-op.
	if (-not [string]::IsNullOrWhiteSpace($connStr) -and $connStr -notmatch '(?i)(^|;)\s*TrustServerCertificate\s*=')
	{
		$connStr = $connStr.TrimEnd().TrimEnd(';') + ';TrustServerCertificate=True;'
	}

	return $connStr
}

# ===================================================================================================
#                              FUNCTION: Download_NSSM
# ---------------------------------------------------------------------------------------------------
# Description:
#   Ensures nssm.exe exists before any function installs or manages services. Can start a
#   hidden worker and return immediately (Background), wait for completion (Wait), or
#   run synchronously. Always publishes the NSSM paths and writes a status log per run.
#   Returns a hashtable @{ Success; Path; Note; Error; Log } describing the outcome.
#
# Parameters:
#   - ToolsRoot            [string]: tools root, defaults to $script:ToolsDir
#   - ForceDownload        [switch]: redownload even if nssm.exe already exists
#   - Background           [switch]: start a hidden worker and return immediately
#   - Wait                 [switch]: block until nssm.exe appears (or timeout)
#   - WaitTimeoutSeconds   [int]: max wait when -Wait is used (default 120)
#   - WaitPollMilliseconds [int]: poll interval when -Wait is used (default 250)
#   - ThrowOnFail          [switch]: throw instead of returning a failure result
#
# Variables:
#   - $script:NssmDir / $script:NssmExe: published NSSM locations
#   - $script:NssmDownloadLog: Download_NSSM_LOG.txt under the script log folder
#
# Writes:
#   - <ToolsRoot>\NSSM\nssm.exe and the status log file
# ===================================================================================================

function Download_NSSM
{
	[CmdletBinding()]
	param (
		# Where tools live. Defaults to the centralized script path configuration.
		[string]$ToolsRoot = $script:ToolsDir,
		# Force redownload even if nssm.exe already exists
		[switch]$ForceDownload,
		# Start download in background and RETURN immediately (non-blocking)
		[switch]$Background,
		# Wait until NSSM exists (blocks only when requested)
		[switch]$Wait,
		# Max wait time when -Wait is used
		[int]$WaitTimeoutSeconds = 120,
		# Poll rate when -Wait is used
		[int]$WaitPollMilliseconds = 250,
		# If you want failures to throw
		[switch]$ThrowOnFail
	)

	# Single log path (overwrite each run)
	$script:NssmDownloadLog = Join-Path $script:ScriptsLog 'Download_NSSM_LOG.txt'

	# ------------------------------------------------------------------------------------------------
	# Target paths
	# ------------------------------------------------------------------------------------------------
	$script:NssmDir = Join-Path $ToolsRoot 'NSSM'
	$script:NssmExe = Join-Path $script:NssmDir 'nssm.exe'

	# Publish paths for other functions
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	$script:FunctionResults['NssmDir'] = $script:NssmDir
	$script:FunctionResults['NssmExe'] = $script:NssmExe

	# Ensure tools folders exist
	try
	{
		if (-not (Test-Path $ToolsRoot)) { New-Item -Path $ToolsRoot -ItemType Directory -Force | Out-Null }
		if (-not (Test-Path $script:NssmDir)) { New-Item -Path $script:NssmDir -ItemType Directory -Force | Out-Null }
	}
	catch
	{
		if ($ThrowOnFail) { throw "Download_NSSM: Failed creating tools folder(s). $($_.Exception.Message)" }
		return @{ Success = $false; Path = $script:NssmExe; Note = 'Tools folder create failed'; Error = $_.Exception.Message; Log = $script:NssmDownloadLog }
	}

	# Fast return if already present and not forcing
	if (-not $ForceDownload -and (Test-Path $script:NssmExe))
	{
		return @{ Success = $true; Path = $script:NssmExe; Note = 'Already present'; Log = $script:NssmDownloadLog }
	}

	# ------------------------------------------------------------------------------------------------
	# Download worker (source: src\Resources\Workers\Download_NSSM_Worker.ps1) - runs inline or in a
	# background runspace; writes only to the status log (overwritten each run)
	# ------------------------------------------------------------------------------------------------
	$downloadScriptText = @'
# ---------------------------------------------------------------------------------------------------
# NSSM download worker - embedded at build time into Download_NSSM ($downloadScriptText).
#   Runs in a 1-thread runspace (Download_NSSM -Background) or inline. Tries the fawno GitHub
#   mirror, then nssm.cc (Invoke-WebRequest -> BITS -> WebClient per URL), extracts, picks the
#   win64/win32 binary, copies it to NssmExe and checks it is > 50 KB.
#   Returns @{ Success; Path; Note; [Error] }; every step goes to LogFile (overwritten per run).
# ---------------------------------------------------------------------------------------------------
param(
	[string]$ToolsRoot,
	[bool]$ForceDownload,
	[string]$NssmDir,
	[string]$NssmExe,
	[string]$LogFile
)

$ErrorActionPreference = "Stop"

# Log writer (scriptblock, not a function)
$logLine = {
	param([string]$Line)
	try
	{
		$ts = (Get-Date -Format "M/d/yyyy h:mm:ss tt")
		("[{0}] {1}" -f $ts, $Line) | Add-Content -Path $LogFile -Encoding UTF8
	}
	catch { }
}

# Overwrite log at start (latest results only)
try { "" | Set-Content -Path $LogFile -Encoding UTF8 -Force } catch { }

& $logLine "Worker start"
& $logLine ("ToolsRoot = {0}" -f $ToolsRoot)
& $logLine ("NssmDir   = {0}" -f $NssmDir)
& $logLine ("NssmExe   = {0}" -f $NssmExe)

try
{
	# Ensure dirs exist
	if (-not (Test-Path $ToolsRoot)) { New-Item -Path $ToolsRoot -ItemType Directory -Force | Out-Null }
	if (-not (Test-Path $NssmDir))   { New-Item -Path $NssmDir   -ItemType Directory -Force | Out-Null }

	# If already present and not forcing
	if (-not $ForceDownload -and (Test-Path $NssmExe))
	{
		& $logLine "Already present. Done."
		return @{ Success = $true; Path = $NssmExe; Note = 'Already present' }
	}

	# TLS 1.2 added to what is enabled (assigning it alone would switch TLS 1.3 off where it exists)
	try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
	# Windows PowerShell 5.1 redraws a progress bar per chunk - it makes Invoke-WebRequest many times slower
	$ProgressPreference = 'SilentlyContinue'

	$is64 = [Environment]::Is64BitOperatingSystem

	# URLs (multiple sources)
	$urls = @()
	if ($is64) { $urls += "https://github.com/fawno/nssm.cc/releases/download/v2.24.1/nssm-v2.24.1-Win64.zip" }
	else       { $urls += "https://github.com/fawno/nssm.cc/releases/download/v2.24.1/nssm-v2.24.1-Win32.zip" }
	$urls += "https://nssm.cc/release/nssm-2.24.zip"
	$urls += "https://nssm.cc/ci/nssm-2.24-101-g897c7ad.zip"

	# Temp workspace
	$tempRoot = Join-Path $env:TEMP ("NSSM_" + [Guid]::NewGuid().ToString("N"))
	$tempZip  = Join-Path $tempRoot "nssm.zip"
	$tempOut  = Join-Path $tempRoot "out"
	New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null
	New-Item -Path $tempOut  -ItemType Directory -Force | Out-Null

	$downloaded = $false
	$lastErr    = ""

	foreach ($u in $urls)
	{
		& $logLine ("Trying URL: {0}" -f $u)

		# wipe old zip
		try { if (Test-Path $tempZip) { Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue } } catch { }

		# Method 1: Invoke-WebRequest (timeout)
		try
		{
			Invoke-WebRequest -Uri $u -OutFile $tempZip -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop

			# IMPORTANT FIX: parentheses so "-and" is an operator
			if ((Test-Path $tempZip) -and ((Get-Item $tempZip).Length -gt 0))
			{
				$downloaded = $true
				& $logLine "Download OK via Invoke-WebRequest"
				break
			}
		}
		catch
		{
			$lastErr = $_.Exception.Message
			& $logLine ("Invoke-WebRequest failed: {0}" -f $lastErr)
		}

		# Method 2: BITS (if available)
		try
		{
			if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue)
			{
				Start-BitsTransfer -Source $u -Destination $tempZip -Priority Foreground -ErrorAction Stop

				# IMPORTANT FIX: parentheses so "-and" is an operator
				if ((Test-Path $tempZip) -and ((Get-Item $tempZip).Length -gt 0))
				{
					$downloaded = $true
					& $logLine "Download OK via BITS"
					break
				}
			}
		}
		catch
		{
			$lastErr = $_.Exception.Message
			& $logLine ("BITS failed: {0}" -f $lastErr)
		}

		# Method 3: WebClient fallback
		try
		{
			$wc = New-Object System.Net.WebClient
			try { $wc.DownloadFile($u, $tempZip) }
			finally { $wc.Dispose() }

			# IMPORTANT FIX: parentheses so "-and" is an operator
			if ((Test-Path $tempZip) -and ((Get-Item $tempZip).Length -gt 0))
			{
				$downloaded = $true
				& $logLine "Download OK via WebClient"
				break
			}
		}
		catch
		{
			$lastErr = $_.Exception.Message
			& $logLine ("WebClient failed: {0}" -f $lastErr)
		}
	}

	if (-not $downloaded)
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine ("RESULT: FAILED - Download failed from all sources. Last error: {0}" -f $lastErr)
		return @{ Success = $false; Path = $NssmExe; Note = 'Download failed'; Error = $lastErr }
	}

	# Unblock zip
	try { Unblock-File -Path $tempZip -ErrorAction SilentlyContinue } catch { }

	# Extract
	$extracted = $false
	try
	{
		Expand-Archive -Path $tempZip -DestinationPath $tempOut -Force -ErrorAction Stop
		$extracted = $true
		& $logLine "Extract OK via Expand-Archive"
	}
	catch
	{
		try
		{
			Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
			[System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $tempOut)
			$extracted = $true
			& $logLine "Extract OK via ZipFile"
		}
		catch
		{
			$extracted = $false
			& $logLine ("Extract failed: {0}" -f $_.Exception.Message)
		}
	}

	if (-not $extracted)
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine "RESULT: FAILED - Extraction failed."
		return @{ Success = $false; Path = $NssmExe; Note = 'Extraction failed' }
	}

	# Find nssm.exe
	$exeCandidates = Get-ChildItem -Path $tempOut -Recurse -Filter "nssm.exe" -File -ErrorAction SilentlyContinue
	if (-not $exeCandidates -or $exeCandidates.Count -eq 0)
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine "RESULT: FAILED - nssm.exe not found in extracted content."
		return @{ Success = $false; Path = $NssmExe; Note = 'nssm.exe not found' }
	}

	# Choose proper binary
	$chosen = $null
	if ($is64)
	{
		$chosen = $exeCandidates | Where-Object { $_.FullName -match "(\\|/)(win64|x64|amd64)(\\|/)" } | Select-Object -First 1
		if (-not $chosen) { $chosen = $exeCandidates | Select-Object -First 1 }
	}
	else
	{
		$chosen = $exeCandidates | Where-Object { $_.FullName -match "(\\|/)(win32|x86|i386)(\\|/)" } | Select-Object -First 1
		if (-not $chosen) { $chosen = $exeCandidates | Select-Object -First 1 }
	}

	if (-not $chosen -or -not (Test-Path $chosen.FullName))
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine "RESULT: FAILED - Could not select a valid nssm.exe candidate."
		return @{ Success = $false; Path = $NssmExe; Note = 'Selection failed' }
	}

	& $logLine ("Selected: {0}" -f $chosen.FullName)

	# Copy to stable location
	Copy-Item -Path $chosen.FullName -Destination $NssmExe -Force -ErrorAction Stop

	# Validate
	if (-not (Test-Path $NssmExe) -or ((Get-Item $NssmExe).Length -lt 50KB))
	{
		try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
		& $logLine ("RESULT: FAILED - Validation failed after copy to '{0}'." -f $NssmExe)
		return @{ Success = $false; Path = $NssmExe; Note = 'Validation failed' }
	}

	# Cleanup
	try { Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }

	& $logLine ("RESULT: SUCCESS - NSSM ready at {0}" -f $NssmExe)
	return @{ Success = $true; Path = $NssmExe; Note = 'Downloaded/Updated' }
}
catch
{
	& $logLine ("RESULT: FAILED - Unhandled exception: {0}" -f $_.Exception.Message)
	return @{ Success = $false; Path = $NssmExe; Note = 'Exception'; Error = $_.Exception.Message }
}
'@

	# ------------------------------------------------------------------------------------------------
	# Background job (in-process runspace) - avoids hidden spawned PowerShell that AV hates
	# ------------------------------------------------------------------------------------------------
	if ($Background)
	{
		# If a job is already running, don't start another
		if ($script:NssmDownloadJob -and $script:NssmDownloadJob.Handle -and (-not $script:NssmDownloadJob.Handle.IsCompleted))
		{
			return @{ Success = $true; Path = $script:NssmExe; Note = 'Already downloading in background'; Log = $script:NssmDownloadLog }
		}
		# A finished earlier download nobody waited for: free its runspace before starting again
		if ($script:NssmDownloadJob)
		{
			try { $script:NssmDownloadJob.PS.Dispose(); $script:NssmDownloadJob.Pool.Close(); $script:NssmDownloadJob.Pool.Dispose() }
			catch { }
			$script:NssmDownloadJob = $null
		}

		try
		{
			# Create a 1-thread runspace pool
			$pool = [RunspaceFactory]::CreateRunspacePool(1, 1)
			$pool.ApartmentState = 'MTA'
			$pool.ThreadOptions = 'ReuseThread'
			$pool.Open()

			$ps = [PowerShell]::Create()
			$ps.RunspacePool = $pool

			$null = $ps.AddScript($downloadScriptText)
			$null = $ps.AddArgument($ToolsRoot)
			$null = $ps.AddArgument([bool]$ForceDownload.IsPresent)
			$null = $ps.AddArgument($script:NssmDir)
			$null = $ps.AddArgument($script:NssmExe)
			$null = $ps.AddArgument($script:NssmDownloadLog)

			$handle = $ps.BeginInvoke()

			# Keep a handle so we can wait later if needed
			$script:NssmDownloadJob = @{
				PS	    = $ps
				Pool    = $pool
				Handle  = $handle
				Started = Get-Date
				Log	    = $script:NssmDownloadLog
				Exe	    = $script:NssmExe
			}

			# Return immediately unless -Wait is also requested
			if (-not $Wait)
			{
				return @{ Success = $true; Path = $script:NssmExe; Note = 'Started background download'; Log = $script:NssmDownloadLog }
			}
		}
		catch
		{
			if ($ThrowOnFail) { throw "Download_NSSM: Failed to start background runspace. $($_.Exception.Message)" }
			return @{ Success = $false; Path = $script:NssmExe; Note = 'Background start failed'; Error = $_.Exception.Message; Log = $script:NssmDownloadLog }
		}
	}

	# ------------------------------------------------------------------------------------------------
	# Wait mode (only blocks when asked). Waiting only makes sense while a background download
	# runs - without one, fall through to the synchronous download below instead of waiting
	# for a file nobody is fetching.
	# ------------------------------------------------------------------------------------------------
	$bgRunning = ($script:NssmDownloadJob -and $script:NssmDownloadJob.Handle -and -not $script:NssmDownloadJob.Handle.IsCompleted)
	if ($Wait -and $bgRunning)
	{
		# Wait for the download itself to finish (not just for the file to appear - it may still
		# be copying), keeping the window responsive.
		$job = $script:NssmDownloadJob
		$deadline = (Get-Date).AddSeconds($WaitTimeoutSeconds)
		while ((Get-Date) -lt $deadline -and -not $job.Handle.IsCompleted)
		{
			try { [System.Windows.Forms.Application]::DoEvents() }
			catch { }
			Start-Sleep -Milliseconds $WaitPollMilliseconds
		}

		if (-not $job.Handle.IsCompleted)
		{
			if ($ThrowOnFail) { throw "Download_NSSM: Timed out waiting for NSSM. Check log: $script:NssmDownloadLog" }
			return @{ Success = $false; Path = $script:NssmExe; Note = 'Timed out waiting for NSSM'; Log = $script:NssmDownloadLog }
		}

		$jobResult = $null
		try { $jobResult = @($job.PS.EndInvoke($job.Handle)) | Select-Object -Last 1 }
		catch { }
		try { $job.PS.Dispose(); $job.Pool.Close(); $job.Pool.Dispose() }
		catch { }
		$script:NssmDownloadJob = $null

		if (Test-Path -LiteralPath $script:NssmExe)
		{
			return @{ Success = $true; Path = $script:NssmExe; Note = 'NSSM ready'; Log = $script:NssmDownloadLog }
		}
		$why = if ($jobResult -and $jobResult.Note) { [string]$jobResult.Note } else { 'Download failed' }
		if ($ThrowOnFail) { throw "Download_NSSM: $why. Check log: $script:NssmDownloadLog" }
		return @{ Success = $false; Path = $script:NssmExe; Note = $why; Log = $script:NssmDownloadLog }
	}

	# ------------------------------------------------------------------------------------------------
	# Synchronous (blocking): no -Background, or -Wait with no background download running
	# ------------------------------------------------------------------------------------------------
	try
	{
		$sb = [ScriptBlock]::Create($downloadScriptText)
		$result = & $sb $ToolsRoot ([bool]$ForceDownload.IsPresent) $script:NssmDir $script:NssmExe $script:NssmDownloadLog

		if ($result -and $result.Success)
		{
			return @{ Success = $true; Path = $script:NssmExe; Note = $result.Note; Log = $script:NssmDownloadLog }
		}

		if ($ThrowOnFail) { throw "Download_NSSM: Failed. Check log: $script:NssmDownloadLog" }
		return @{ Success = $false; Path = $script:NssmExe; Note = 'Failed'; Log = $script:NssmDownloadLog }
	}
	catch
	{
		if ($ThrowOnFail) { throw "Download_NSSM: Exception. $($_.Exception.Message). Log: $script:NssmDownloadLog" }
		return @{ Success = $false; Path = $script:NssmExe; Note = 'Exception'; Error = $_.Exception.Message; Log = $script:NssmDownloadLog }
	}
}

# ===================================================================================================
#                              FUNCTION: Repair_BMS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Repairs the Bizerba BMS service on the local machine: disables the service, stops it with
#   escalating actions (Stop-Service, sc stop, net stop /y, WMI, taskkill by PID/service/image),
#   cleans the toBizerba staging and terminals cache folders, deletes and re-registers the
#   service (BMSSrv.exe -reg) and starts it. Requires elevation.
#
# Parameters:
#   - BMSSrvPath         [string]: path to BMSSrv.exe (default $script:BmssrvExePath)
#   - CleanToBizerba     [bool]: purge the toBizerba staging folder (default $true)
#   - RemoveTerminalsCache [bool]: remove the terminals cache folder (default $true)
#
# Variables:
#   - $script:BmssrvExePath: default BMSSrv.exe location
#   - $script:BmsRootPath: BMS install root (toBizerba/terminals folders)
#
# Writes:
#   - local BMS service state, toBizerba and terminals folders
# ===================================================================================================

function Repair_BMS
{
	[CmdletBinding()]
	param
	(
		# Full path to BMSSrv.exe (used to re-register the service)
		[Parameter(Mandatory = $false)]
		[string]$BMSSrvPath = $script:BmssrvExePath,
		# Purge staging folder (matches your fast .bat)
		[Parameter(Mandatory = $false)]
		[bool]$CleanToBizerba = $true,
		# Remove terminals cache folder (matches your fast .bat)
		[Parameter(Mandatory = $false)]
		[bool]$RemoveTerminalsCache = $true
	)

	Write_Log "`r`n==================== Starting Repair_BMS (IMMEDIATE TASKKILL) ====================`r`n" "blue"

	# ------------------------------------------------------------------------------------------------
	# 0) Elevation check (required for service control + protected folder deletes)
	# ------------------------------------------------------------------------------------------------
	$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
	IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	if (-not $isAdmin)
	{
		Write_Log "Insufficient permissions. Please run as Administrator." "red"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# 1) Validate BMSSrv.exe path
	# ------------------------------------------------------------------------------------------------
	if (-not (Test-Path -LiteralPath $BMSSrvPath))
	{
		Write_Log "BMSSrv.exe not found at: $BMSSrvPath" "red"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# 2) Constants / paths
	# ------------------------------------------------------------------------------------------------
	$serviceName = "BMS"
	$bmsRoot = $script:BmsRootPath
	$toBizerbaPath = Join-Path $bmsRoot "toBizerba"
	$terminalsPath = Join-Path $bmsRoot "terminals"

	# Kill only what you want (silent)
	$killImages = @(
		"BMSSrv.exe",
		"BMS.exe"
	)

	Write_Log "Phase 1: IMMEDIATE kill (no wait loops)..." "blue"

	# ------------------------------------------------------------------------------------------------
	# 3) Kill from the get-go (NO waiting) - SILENT
	# ------------------------------------------------------------------------------------------------
	try
	{
		cmd /c 'taskkill /F /T /FI "SERVICES eq BMS" >nul 2>&1'
		Write_Log 'Issued: taskkill /F /T /FI "SERVICES eq BMS" (silent)' "yellow"
	}
	catch
	{
		Write_Log "taskkill by service filter threw an exception (often harmless): $_" "yellow"
	}

	foreach ($img in $killImages)
	{
		try
		{
			cmd /c ("taskkill /F /T /IM {0} >nul 2>&1" -f $img)
			Write_Log "Issued: taskkill /F /T /IM $img (silent)" "yellow"
		}
		catch
		{
			Write_Log "taskkill for $img threw an exception (often harmless): $_" "yellow"
		}
	}

	# ------------------------------------------------------------------------------------------------
	# 4) Disable/Stop service quickly (best-effort, no waits)
	#    - Skip noisy net stop unless service is actually not stopped
	# ------------------------------------------------------------------------------------------------
	Write_Log "Phase 2: Disable/Stop service (best-effort, no waits)..." "blue"

	$svcBefore = $null
	try { $svcBefore = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcBefore = $null }

	try { cmd /c ("sc.exe config {0} start= disabled >nul 2>&1" -f $serviceName) }
	catch { }
	try { Set-Service -Name $serviceName -StartupType Disabled -ErrorAction SilentlyContinue }
	catch { }

	if ($svcBefore -and $svcBefore.Status -ne 'Stopped')
	{
		try { cmd /c ("sc.exe stop {0} >nul 2>&1" -f $serviceName) }
		catch { }
		try { cmd /c ("net stop {0} /y >nul 2>&1" -f $serviceName) }
		catch { } # silent
		try { Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue }
		catch { }

		# Re-kill once more (still no waits)
		try { cmd /c 'taskkill /F /T /FI "SERVICES eq BMS" >nul 2>&1' }
		catch { }
		foreach ($img in $killImages)
		{
			try { cmd /c ("taskkill /F /T /IM {0} >nul 2>&1" -f $img) }
			catch { }
		}
	}
	else
	{
		Write_Log "Service is already stopped or not present; skipping net stop to avoid console noise." "gray"
	}

	# ------------------------------------------------------------------------------------------------
	# 5) Cleanup filesystem (fast)
	# ------------------------------------------------------------------------------------------------
	Write_Log "Phase 3: Cleanup folders..." "blue"

	if ($CleanToBizerba)
	{
		if (Test-Path -LiteralPath $toBizerbaPath)
		{
			try
			{
				Get-ChildItem -LiteralPath $toBizerbaPath -Force -ErrorAction SilentlyContinue |
				Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

				Write_Log "Cleared: $toBizerbaPath" "green"
			}
			catch
			{
				Write_Log "Failed clearing ${toBizerbaPath}: $_" "yellow"
			}
		}
		else
		{
			Write_Log "Folder not found (skip): $toBizerbaPath" "gray"
		}
	}
	else
	{
		Write_Log "CleanToBizerba disabled (skip)." "gray"
	}

	if ($RemoveTerminalsCache)
	{
		if (Test-Path -LiteralPath $terminalsPath)
		{
			try
			{
				Remove-Item -LiteralPath $terminalsPath -Recurse -Force -ErrorAction SilentlyContinue
				Write_Log "Removed cache folder: $terminalsPath" "green"
			}
			catch
			{
				Write_Log "Failed removing ${terminalsPath}: $_" "yellow"
			}
		}
		else
		{
			Write_Log "Folder not found (skip): $terminalsPath" "gray"
		}
	}
	else
	{
		Write_Log "RemoveTerminalsCache disabled (skip)." "gray"
	}

	# ------------------------------------------------------------------------------------------------
	# 6) Unregister (delete service) + Register again (BMSSrv.exe -reg)
	#    - CHANGED: BMSSrv.exe output is redirected so it doesn't print to console
	# ------------------------------------------------------------------------------------------------
	Write_Log "Phase 4: Unregister (sc delete) + Register (BMSSrv.exe -reg)..." "blue"

	try
	{
		# Pre-kill again (helps with "marked for deletion" locks)
		try { cmd /c 'taskkill /F /T /FI "SERVICES eq BMS" >nul 2>&1' }
		catch { }
		foreach ($img in $killImages)
		{
			try { cmd /c ("taskkill /F /T /IM {0} >nul 2>&1" -f $img) }
			catch { }
		}

		cmd /c ("sc.exe delete {0} >nul 2>&1" -f $serviceName)
		Write_Log "Issued: sc delete $serviceName (silent)" "yellow"
	}
	catch
	{
		Write_Log "sc delete threw an exception (service may already be gone): $_" "yellow"
	}

	# --- Register service again, SILENT (no console spam) ---
	$tmpOut = Join-Path $env:TEMP ("BMSSrv_reg_stdout_{0}.log" -f ([Guid]::NewGuid().ToString("N")))
	$tmpErr = Join-Path $env:TEMP ("BMSSrv_reg_stderr_{0}.log" -f ([Guid]::NewGuid().ToString("N")))

	try
	{
		# Ensure temp files don't exist
		try { if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue } }
		catch { }
		try { if (Test-Path -LiteralPath $tmpErr) { Remove-Item -LiteralPath $tmpErr -Force -ErrorAction SilentlyContinue } }
		catch { }

		# CHANGED: Redirect output + hide window so BMSSrv does not write to the console
		$proc = Start-Process -FilePath $BMSSrvPath -ArgumentList "-reg" -WindowStyle Hidden -Wait -PassThru `
							  -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr

		if ($proc.ExitCode -ne 0)
		{
			$stdout = ""; $stderr = ""
			try { if (Test-Path -LiteralPath $tmpOut) { $stdout = (Get-Content -LiteralPath $tmpOut -ErrorAction SilentlyContinue | Out-String).Trim() } }
			catch { }
			try { if (Test-Path -LiteralPath $tmpErr) { $stderr = (Get-Content -LiteralPath $tmpErr -ErrorAction SilentlyContinue | Out-String).Trim() } }
			catch { }

			Write_Log ("BMSSrv.exe -reg failed (exit code {0})." -f $proc.ExitCode) "red"
			if ($stdout) { Write_Log ("BMSSrv stdout: {0}" -f $stdout) "yellow" }
			if ($stderr) { Write_Log ("BMSSrv stderr: {0}" -f $stderr) "yellow" }
			return
		}

		Write_Log "BMSSrv.exe -reg completed successfully (silent)." "green"
	}
	catch
	{
		Write_Log "BMSSrv.exe -reg failed: $_" "red"
		return
	}
	finally
	{
		# Cleanup redirected output files (keep it tidy)
		try { if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force -ErrorAction SilentlyContinue } }
		catch { }
		try { if (Test-Path -LiteralPath $tmpErr) { Remove-Item -LiteralPath $tmpErr -Force -ErrorAction SilentlyContinue } }
		catch { }
	}

	# ------------------------------------------------------------------------------------------------
	# 7) Re-enable + start service (avoid "service is not started" noise)
	# ------------------------------------------------------------------------------------------------
	Write_Log "Phase 5: Start service..." "blue"

	try { cmd /c ("sc.exe config {0} start= auto >nul 2>&1" -f $serviceName) }
	catch { }
	try { Set-Service -Name $serviceName -StartupType Automatic -ErrorAction SilentlyContinue }
	catch { }

	$svcAfterReg = $null
	try { $svcAfterReg = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcAfterReg = $null }

	if ($svcAfterReg -and $svcAfterReg.Status -ne 'Running')
	{
		try { cmd /c ("net start {0} /y >nul 2>&1" -f $serviceName) }
		catch { }
		try { Start-Service -Name $serviceName -ErrorAction SilentlyContinue }
		catch { }
	}
	else
	{
		Write_Log "Service already running or not present immediately after registration; skipping net start." "gray"
	}

	$svcFinal = $null
	try { $svcFinal = Get-Service -Name $serviceName -ErrorAction Stop }
	catch { $svcFinal = $null }

	if ($svcFinal -and $svcFinal.Status -eq 'Running')
	{
		Write_Log "'$serviceName' is running." "green"
	}
	else
	{
		$st = if ($svcFinal) { $svcFinal.Status }
		else { "NotFound" }
		Write_Log "WARNING: Could not confirm '$serviceName' running. Current state: $st" "yellow"
	}

	Write_Log "`r`n==================== Repair_BMS Completed (IMMEDIATE TASKKILL) ====================`r`n" "blue"
}

# ===================================================================================================
#                              FUNCTION: Remove_Duplicate_Files_From_toBizerba
# ---------------------------------------------------------------------------------------------------
# Description:
#   Duplicate File Monitor for C:\Bizerba\RetailConnect\BMS\toBizerba: prompt to run it now
#   (background job, controllable from the GUI) or install it as an always-on Windows service.
#   Removes duplicates by content hash (keeps the oldest by CreationTime) and also removes
#   files whose non-empty lines are fully contained in another file.
#
# Parameters:
#   (none)
#
# Variables:
#   - $script:BmsToBizerbaPath: folder being monitored
#   - $script:DuplicateFileMonitorServiceRoot: service root (ProgramData); watcher script goes in its Scripts subfolder
#   - $script:ScriptsLog: single combined log (service script + NSSM stdout/stderr)
#   - $script:FunctionResults['NssmExe'] / $script:NssmExe / $script:NssmDefaultPath: NSSM candidates
#
# Notes:
#   - Never downloads NSSM; if nssm.exe is missing, install/uninstall errors out and stops.
# ===================================================================================================

function Remove_Duplicate_Files_From_toBizerba
{
	[CmdletBinding()]
	param ()

	Write_Log "`r`n==================== Starting Remove_Duplicate_Files_From_toBizerba Function ====================`r`n" "blue"

	# ==========================================================================================
	# CORE PATHS
	# ==========================================================================================
	$TargetPath = $script:BmsToBizerbaPath
	$serviceName = "Remove_Duplicate_Files_From_toBizerba"

	$serviceRoot = $script:DuplicateFileMonitorServiceRoot
	$svcScripts = Join-Path $serviceRoot "Scripts"

	$psScriptName = "Remove_Duplicate_Files_From_toBizerba.ps1"
	$psScriptPath = Join-Path $svcScripts $psScriptName

	# CHANGE: single combined log for everything (service script + NSSM stdout/stderr)
	$singleLog = Join-Path $script:ScriptsLog "Remove_Duplicate_Files_From_toBizerba.log"

	# ==========================================================================================
	# Ensure folders exist
	# ==========================================================================================
	foreach ($p in @($serviceRoot, $svcScripts))
	{
		if (-not (Test-Path -LiteralPath $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null }
	}

	# ==========================================================================================
	# Resolve NSSM executable (NO downloads)
	# ==========================================================================================
	$nssmExe = $null
	$nssmCandidates = @()

	try
	{
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('NssmExe') -and $script:FunctionResults['NssmExe'])
		{
			$nssmCandidates += [string]$script:FunctionResults['NssmExe']
		}
	}
	catch { }

	try
	{
		if ($script:NssmExe) { $nssmCandidates += [string]$script:NssmExe }
	}
	catch { }

	$nssmCandidates += $script:NssmDefaultPath

	foreach ($c in $nssmCandidates)
	{
		if ([string]::IsNullOrWhiteSpace($c)) { continue }
		if (Test-Path -LiteralPath $c) { $nssmExe = $c; break }
	}

	# ==========================================================================================
	# WRITE WATCHER SCRIPT (NO interpolation bugs; concise logging; never exits)
	# ==========================================================================================
	$psScriptContent = @'
#requires -Version 5.1
param([int]$IntervalSeconds = 5)

$ErrorActionPreference = 'Continue'

$Path    = '__TARGETPATH__'
$LogFile = '__LOGFILE__'

# Ensure log dir exists
try {
	$ld = Split-Path -Parent $LogFile
	if (-not (Test-Path -LiteralPath $ld)) { New-Item -Path $ld -ItemType Directory -Force | Out-Null }
} catch { }

# Concise log helper (scriptblock, not a function)
$log = {
	param([string]$m)
	try { Add-Content -LiteralPath $LogFile -Value ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $m) -Encoding UTF8 } catch { }
}

# Ensure target folder exists (don't exit if missing)
try {
	if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
} catch { }

& $log ("Service script started. Interval={0}s Path={1}" -f $IntervalSeconds, $Path)

while ($true)
{
	try
	{
		# Re-create folder if deleted
		if (-not (Test-Path -LiteralPath $Path))
		{
			try { New-Item -Path $Path -ItemType Directory -Force | Out-Null } catch { }
			Start-Sleep -Seconds $IntervalSeconds
			continue
		}

		$removedHash = 0
		$removedLine = 0
		$removedSamples = New-Object System.Collections.Generic.List[string]

		# --------------------------
		# 1) Remove duplicates by SHA256 hash (keep oldest by CreationTime)
		# --------------------------
		$files = Get-ChildItem -LiteralPath $Path -File -ErrorAction SilentlyContinue
		$hashTable = @{}

		foreach ($file in $files)
		{
			try { if ($file.Length -eq 0) { continue } } catch { }

			try { $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash }
			catch { continue }

			if (-not $hashTable.ContainsKey($hash)) { $hashTable[$hash] = @() }
			$hashTable[$hash] += $file
		}

		foreach ($entry in $hashTable.GetEnumerator())
		{
			$list = $entry.Value
			if ($list.Count -gt 1)
			{
				$list = $list | Sort-Object CreationTime
				$keep = $list[0]
				$dups = $list[1..($list.Count - 1)]

				foreach ($d in $dups)
				{
					try
					{
						Remove-Item -LiteralPath $d.FullName -Force -ErrorAction Stop
						$removedHash++
						if ($removedSamples.Count -lt 5) { [void]$removedSamples.Add(("HASH:{0}" -f $d.Name)) }
					}
					catch { }
				}
			}
		}

		# --------------------------
		# 2) Remove by line containment: a file whose non-empty lines all appear in another file
		#    (with at least as many lines) goes. Each file is read ONCE per pass (it was re-read
		#    for every pair and the scan restarted after each delete).
		# --------------------------
		$files = @(Get-ChildItem -LiteralPath $Path -File -ErrorAction SilentlyContinue)
		$lineSets = @{ }
		$lineCounts = @{ }
		foreach ($fi in $files)
		{
			$lines = @(Get-Content -LiteralPath $fi.FullName -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim() -ne "" })
			if ($lines.Count -eq 0) { continue }
			$set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
			foreach ($ln in $lines) { [void]$set.Add($ln.Trim()) }
			$lineSets[$fi.FullName] = $set
			$lineCounts[$fi.FullName] = $lines.Count
		}

		$deleted = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
		foreach ($fileA in $files)
		{
			$a = $fileA.FullName
			if (-not $lineSets.ContainsKey($a)) { continue }
			foreach ($fileB in $files)
			{
				$b = $fileB.FullName
				if ($a -eq $b -or $deleted.Contains($b) -or -not $lineSets.ContainsKey($b)) { continue }
				if ($lineCounts[$b] -lt $lineCounts[$a]) { continue }
				if (-not $lineSets[$a].IsSubsetOf($lineSets[$b])) { continue }

				# Only delete when nobody has the file open (BMS may still be writing it)
				$canDelete = $true
				try
				{
					$s = [System.IO.File]::Open($a, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
					$s.Close()
				}
				catch { $canDelete = $false }

				if ($canDelete)
				{
					try
					{
						Remove-Item -LiteralPath $a -Force -ErrorAction Stop
						[void]$deleted.Add($a)
						$removedLine++
						if ($removedSamples.Count -lt 5) { [void]$removedSamples.Add(("LINES:{0}" -f $fileA.Name)) }
					}
					catch { }
				}
				break
			}
		}
		# Concise: only log when something was removed
		if (($removedHash + $removedLine) -gt 0)
		{
			$sampleText = ""
			if ($removedSamples.Count -gt 0) { $sampleText = " Samples: " + ($removedSamples -join ", ") }
			& $log ("Removed: Hash={0}, Lines={1}.{2}" -f $removedHash, $removedLine, $sampleText)
		}
	}
	catch
	{
		& $log ("ERROR: " + $_.Exception.Message)
	}

	Start-Sleep -Seconds $IntervalSeconds
}
'@

	# Inject paths safely into the single-quoted script (escape single quotes for PS literals)
	$tpEsc = $TargetPath.Replace("'", "''")
	$lfEsc = $singleLog.Replace("'", "''")

	$psScriptContent = $psScriptContent.Replace('__TARGETPATH__', $tpEsc)
	$psScriptContent = $psScriptContent.Replace('__LOGFILE__', $lfEsc)

	Set-Content -LiteralPath $psScriptPath -Value $psScriptContent -Encoding UTF8 -Force
	if (-not (Test-Path -LiteralPath $psScriptPath))
	{
		Write_Log "ERROR: Failed to write watcher script: $psScriptPath" "red"
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}

	# ==========================================================================================
	# SERVICE EXISTS?
	# ==========================================================================================
	$hasService = [bool](Get-Service -Name $serviceName -ErrorAction SilentlyContinue)

	# ==========================================================================================
	# GUI
	# ==========================================================================================
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Duplicate File Monitor"
	$form.Size = New-Object System.Drawing.Size(440, 230)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = 'FixedDialog'
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	$label = New-Object System.Windows.Forms.Label
	$label.Text = "How do you want to run the Duplicate File Monitor for:`r`n$TargetPath"
	$label.Location = New-Object System.Drawing.Point(16, 15)
	$label.Size = New-Object System.Drawing.Size(390, 40)
	$label.TextAlign = 'MiddleCenter'
	$form.Controls.Add($label)

	$btnRunNow = New-Object System.Windows.Forms.Button
	$btnRunNow.Text = "Run Now (as Job)"
	$btnRunNow.Location = New-Object System.Drawing.Point(10, 70)
	$btnRunNow.Size = New-Object System.Drawing.Size(110, 35)
	$form.Controls.Add($btnRunNow)

	$btnService = New-Object System.Windows.Forms.Button
	$btnService.Text = "Install Service"
	$btnService.Location = New-Object System.Drawing.Point(125, 70)
	$btnService.Size = New-Object System.Drawing.Size(140, 35)
	$form.Controls.Add($btnService)

	$btnDeleteService = New-Object System.Windows.Forms.Button
	$btnDeleteService.Text = "Uninstall Service"
	$btnDeleteService.Location = New-Object System.Drawing.Point(270, 70)
	$btnDeleteService.Size = New-Object System.Drawing.Size(140, 35)
	$btnDeleteService.Enabled = $hasService
	$form.Controls.Add($btnDeleteService)

	$lblSec = New-Object System.Windows.Forms.Label
	$lblSec.Text = "Interval (seconds):"
	$lblSec.Location = New-Object System.Drawing.Point(70, 120)
	$lblSec.Size = New-Object System.Drawing.Size(110, 22)
	$form.Controls.Add($lblSec)

	$numSec = New-Object System.Windows.Forms.NumericUpDown
	$numSec.Location = New-Object System.Drawing.Point(185, 118)
	$numSec.Size = New-Object System.Drawing.Size(50, 24)
	$numSec.Minimum = 1
	$numSec.Maximum = 3600
	$numSec.Value = 5
	$form.Controls.Add($numSec)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(10, 148)
	$btnCancel.Size = New-Object System.Drawing.Size(400, 35)
	$form.Controls.Add($btnCancel)

	$state = @{ Action = $null; Interval = 5 }
	$state.Job = $null # handlers assign through $state: a plain $monitorJob = ... inside a handler is a new local

	$form.Add_FormClosing({
			if ($state.Job -and $state.Job.State -eq 'Running')
			{
				try { Stop-Job -Job $state.Job; Remove-Job -Job $state.Job -Force; Write_Log "Monitor job stopped on form close." "yellow" }
				catch { }
			}

			Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
				$_.CommandLine -and $_.CommandLine -match [Regex]::Escape($psScriptPath)
			} | ForEach-Object {
				try { Stop-Process -Id $_.ProcessId -Force; Write_Log "Stopped powershell.exe PID $($_.ProcessId) on form close" "yellow" }
				catch { }
			}
		})

	$btnRunNow.Add_Click({
			if ($btnRunNow.Text -eq "Run Now (as Job)")
			{
				$state.Interval = [int]$numSec.Value
				try { "" | Set-Content -LiteralPath $singleLog -Encoding UTF8 -Force }
				catch { }

				$state.Job = Start-Job -Name 'DuplicateFileMonitorJob' -ScriptBlock {
					param ($scriptPath,
						$interval,
						$logFile)
					& "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $scriptPath -IntervalSeconds $interval *>> $logFile
				} -ArgumentList $psScriptPath, $state.Interval, $singleLog

				Write_Log "Started monitor job (ID: $($state.Job.Id)). Log: $singleLog" "green"

				$btnRunNow.Text = "Stop Monitor"
				$btnService.Enabled = $false
				$btnDeleteService.Enabled = $false
				$btnCancel.Text = "Close (stops the monitor)"
			}
			else
			{
				if ($state.Job -and $state.Job.State -eq 'Running')
				{
					try { Stop-Job -Job $state.Job; Remove-Job -Job $state.Job -Force; Write_Log "Monitor job stopped by user." "yellow" }
					catch { }
				}

				Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object {
					$_.CommandLine -and $_.CommandLine -match [Regex]::Escape($psScriptPath)
				} | ForEach-Object {
					try { Stop-Process -Id $_.ProcessId -Force; Write_Log "Stopped powershell.exe PID $($_.ProcessId) by user stop" "yellow" }
					catch { }
				}

				$btnRunNow.Text = "Run Now (as Job)"
				$btnService.Enabled = $true
				$btnDeleteService.Enabled = $hasService
				$btnCancel.Text = "Cancel"
			}
		})

	$btnService.Add_Click({ $state.Interval = [int]$numSec.Value; $state.Action = "service"; $form.Close() })
	$btnDeleteService.Add_Click({ $state.Action = "uninstall"; $form.Close() })
	$btnCancel.Add_Click({ $form.Close() })

	$form.ShowDialog() | Out-Null

	if (-not $state.Action)
	{
		Write_Log "User cancelled Duplicate File Monitor dialog." "yellow"
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}

	# ==========================================================================================
	# ADMIN CHECK FOR SERVICE ACTIONS
	# ==========================================================================================
	if ($state.Action -in @("service", "uninstall"))
	{
		$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
		).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

		if (-not $IsAdmin)
		{
			Write_Log "Service install/uninstall requires Administrator. Re-run as Admin." "red"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}
	}

	# ==========================================================================================
	# UNINSTALL (SC.EXE ONLY)
	# ==========================================================================================
	if ($state.Action -eq "uninstall")
	{
		if (-not (Get-Service -Name $serviceName -ErrorAction SilentlyContinue))
		{
			Write_Log "Service '$serviceName' is not installed." "yellow"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}

		Write_Log "Stopping service '$serviceName' before uninstall..." "yellow"
		try { & $env:ComSpec /c ("sc stop `"{0}`"" -f $serviceName) *> $null }
		catch { }
		try { Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue }
		catch { }
		try { & $env:ComSpec /c ("sc delete `"{0}`"" -f $serviceName) *> $null }
		catch { }

		Start-Sleep -Milliseconds 800
		if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
		{
			Write_Log "ERROR: Service still exists after sc delete. Remove manually in Services.msc." "red"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}

		Write_Log "Service '$serviceName' uninstalled successfully." "green"
		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}

	# ==========================================================================================
	# INSTALL SERVICE (requires NSSM present - NO downloads)
	# ==========================================================================================
	if ($state.Action -eq "service")
	{
		if (-not $nssmExe -or -not (Test-Path -LiteralPath $nssmExe))
		{
			Write_Log "ERROR: NSSM is missing. This function will not download it." "red"
			Write_Log "Expected: $script:NssmDefaultPath" "yellow"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}

		# Fresh log for easy debugging
		try { "" | Set-Content -LiteralPath $singleLog -Encoding UTF8 -Force }
		catch { }

		# Remove existing service (clean reinstall)
		if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
		{
			try { & $env:ComSpec /c ("sc stop `"{0}`"" -f $serviceName) *> $null }
			catch { }
			try { Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue }
			catch { }
			try { & $env:ComSpec /c ("sc delete `"{0}`"" -f $serviceName) *> $null }
			catch { }
			Start-Sleep -Seconds 1
		}

		$psExe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
		$interval = [int]$state.Interval

		# Install with args as separate tokens (no quoting issues)
		& $nssmExe install $serviceName $psExe `
				   '-NoProfile' `
				   '-ExecutionPolicy' 'Bypass' `
				   '-File' $psScriptPath `
				   '-IntervalSeconds' "$interval"

		if ($LASTEXITCODE -ne 0)
		{
			Write_Log "ERROR: NSSM install failed (ExitCode=$LASTEXITCODE). NSSM: $nssmExe" "red"
			Write_Log "Log: $singleLog" "yellow"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}

		# Configure service
		& $nssmExe set $serviceName AppDirectory $svcScripts *> $null
		& $nssmExe set $serviceName Start SERVICE_AUTO_START *> $null
		& $nssmExe set $serviceName ObjectName "LocalSystem" *> $null

		# ONE log file for NSSM stdout + stderr
		& $nssmExe set $serviceName AppStdout $singleLog *> $null
		& $nssmExe set $serviceName AppStderr $singleLog *> $null
		& $nssmExe set $serviceName AppRotateFiles 0 *> $null

		& $nssmExe start $serviceName *> $null
		if ($LASTEXITCODE -ne 0)
		{
			Write_Log "ERROR: Service installed but failed to start (ExitCode=$LASTEXITCODE)." "red"
			Write_Log "Log: $singleLog" "yellow"
			Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
			return
		}

		Start-Sleep -Milliseconds 800
		$s = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
		if ($s -and $s.Status -ne 'Running')
		{
			Write_Log "WARNING: Service not Running (Status=$($s.Status)). Check log: $singleLog" "yellow"
		}
		else
		{
			Write_Log "Service '$serviceName' installed and started. Interval: $interval sec." "green"
			Write_Log "Single log: $singleLog" "green"
			Write_Log "NSSM used: $nssmExe" "gray"
		}

		Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
		return
	}

	Write_Log "No valid action selected. Exiting." "yellow"
	Write_Log "`r`n==================== Remove_Duplicate_Files_From_toBizerba Completed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Edit_TBS_SCL_ver520_Table
# ---------------------------------------------------------------------------------------------------
# Description:
#   Interactive grid editor for the TBS_SCL_ver520 scale table (no row dragging). Auto Sort
#   orders BIZERBA first (by numeric IP) then ISHIDA, renumbers them, applies the naming and
#   BufferTime templates and saves. Save validates and writes exactly what is visible in the
#   grid; Refresh reloads from the DB. IPNetwork must end with a trailing period.
#
# Variables:
#   - $script:FunctionResults: DBSERVER, DBNAME, ConnectionString, SQL_UserID/SQL_UserPwd
#
# Writes:
#   - store database table TBS_SCL_ver520
#
# Notes:
#   - SQL goes through System.Data.SqlClient (no SQL module needed).
#   - Rows are matched on their ORIGINAL brand / model / IPNetwork / IPDevice (NULLs match '').
# ===================================================================================================

function Edit_TBS_SCL_ver520_Table
{
	[CmdletBinding()]
	param ()

	Write_Log "`r`n==================== Starting Edit_TBS_SCL_ver520_Table (No Drag / No Checkbox) ====================`r`n" "blue"

	# ----------------------------------------------------------------------------------------------
	# 1) Resolve DB info and build connection string (SQL Auth preferred, else Integrated)
	# ----------------------------------------------------------------------------------------------
	$server = $script:FunctionResults['DBSERVER']
	$database = $script:FunctionResults['DBNAME']
	$connectionString = Normalize_Sql_ConnectionString -RawConnectionString $script:FunctionResults['ConnectionString'] -Server $server -Database $database -SqlUser $script:FunctionResults['SQL_UserID'] -SqlPwd $script:FunctionResults['SQL_UserPwd']

	if (-not $server -or -not $database)
	{
		Write_Log "DB server or DB name missing - cannot proceed." "red"
		return
	}

	if (-not $connectionString)
	{
		Write_Log "Unable to build a valid SQL connection string from cached values." "red"
		return
	}
	$script:FunctionResults['ConnectionString'] = $connectionString

	# ----------------------------------------------------------------------------------------------
	# 2) SQL helpers (one connection per call)
	# ----------------------------------------------------------------------------------------------
	# Rows of a query (PowerShell unrolls the DataTable: DataRow(s) or nothing)
	$runQuery = {
		param ([string]$Sql)
		$conn = New-Object System.Data.SqlClient.SqlConnection $connectionString
		try
		{
			$conn.Open()
			$cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 0
			$td = New-Object System.Data.DataTable
			[void](New-Object System.Data.SqlClient.SqlDataAdapter $cmd).Fill($td)
			$td
		}
		finally { $conn.Dispose() }
	}

	$runScalar = {
		param ([string]$Sql)
		$conn = New-Object System.Data.SqlClient.SqlConnection $connectionString
		try
		{
			$conn.Open()
			$cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 0
			$cmd.ExecuteScalar()
		}
		finally { $conn.Dispose() }
	}

	# ----------------------------------------------------------------------------------------------
	# 3) Discover the /24 prefix in use; prefer NIC that can reach the SQL server; else NIC w/ gateway; else first RFC1918
	#     Result: $allowedPrefix like '192.168.5' (used to validate/normalize IPNetwork).
	# ----------------------------------------------------------------------------------------------
	$allowedPrefix = $null
	try
	{
		# Resolve SQL server IPv4s (best effort; instance names may not resolve)
		$serverIPv4s = @()
		try
		{
			$serverIPs = [System.Net.Dns]::GetHostAddresses((($server -replace '^(?i)tcp:', '') -split '[\\,]')[0])
			foreach ($ip in $serverIPs)
			{
				if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork)
				{
					$serverIPv4s += $ip
				}
			}
		}
		catch { }

		$allNics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object {
			$_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up
		}

		# CHANGE: Inline subnet comparison (no helper function). Compare (addr & mask) for equality.
		$pickedIP = $null

		# Pass 1: NIC whose IPv4 shares subnet with SQL server
		if ($serverIPv4s.Count -gt 0)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					if ($ua.IPv4Mask -eq $null) { continue }

					$aBytes = $ua.Address.GetAddressBytes()
					$mBytes = $ua.IPv4Mask.GetAddressBytes()

					$matchFound = $false
					foreach ($srv in $serverIPv4s)
					{
						$bBytes = $srv.GetAddressBytes()
						$same = $true
						for ($i = 0; $i -lt 4; $i++)
						{
							if (($aBytes[$i] -band $mBytes[$i]) -ne ($bBytes[$i] -band $mBytes[$i]))
							{
								$same = $false
								break
							}
						}
						if ($same) { $matchFound = $true; break }
					}
					if ($matchFound) { $pickedIP = $ua.Address.ToString(); break }
				}
				if ($pickedIP) { break }
			}
		}

		# Pass 2: NIC with default gateway
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				$hasGw = $false
				foreach ($gw in $ipProps.GatewayAddresses)
				{
					if ($gw.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $hasGw = $true; break }
				}
				if (-not $hasGw) { continue }

				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$pickedIP = $ua.Address.ToString()
					break
				}
				if ($pickedIP) { break }
			}
		}

		# Pass 3: first RFC1918 IPv4
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$ip = $ua.Address.ToString()
					if ($ip.StartsWith('10.') -or $ip.StartsWith('192.168.') -or (
							$ip.StartsWith('172.') -and (([int]($ip.Split('.')[1])) -ge 16) -and (([int]($ip.Split('.')[1])) -le 31)
						))
					{
						$pickedIP = $ip
						break
					}
				}
				if ($pickedIP) { break }
			}
		}

		if ($pickedIP)
		{
			$parts = $pickedIP.Split('.')
			if ($parts.Length -ge 3)
			{
				$allowedPrefix = "$($parts[0]).$($parts[1]).$($parts[2])"
			}
		}
	}
	catch { }

	if ($allowedPrefix)
	{
		Write_Log "IPNetwork allowed prefix detected (active NIC): $allowedPrefix" "cyan"
	}
	else
	{
		Write_Log "WARNING: Could not determine active NIC / RFC1918 IPv4. IPNetwork validation will be skipped (but trailing '.' will still be enforced if parsable)." "yellow" # CHANGE: note trailing '.' enforcement still happens
	}

	# ----------------------------------------------------------------------------------------------
	# 4) Load table -> DataTable (include Orig* columns to anchor updates)
	# ----------------------------------------------------------------------------------------------
	$selectQuery = @"
SELECT
    ScaleCode, ScaleName, ScaleBrand, ScaleModel, IPNetwork, IPDevice, BufferTime, Active
FROM [TBS_SCL_ver520]
ORDER BY ScaleCode ASC;
"@

	$dt = New-Object System.Data.DataTable
	foreach ($c in 'ScaleCode', 'ScaleName', 'ScaleBrand', 'ScaleModel', 'IPNetwork', 'IPDevice', 'BufferTime', 'Active',
		'OrigScaleBrand', 'OrigScaleModel', 'OrigIPNetwork', 'OrigIPDevice')
	{
		[void]$dt.Columns.Add($c)
	}

	# DB rows -> grid rows, with the original key columns snapshotted for the UPDATE match
	$fillGrid = {
		param ($Rows)
		$dt.Rows.Clear()
		foreach ($r in @($Rows))
		{
			if (-not $r) { continue }
			$nr = $dt.NewRow()
			foreach ($col in 'ScaleCode', 'ScaleName', 'ScaleBrand', 'ScaleModel', 'IPNetwork', 'IPDevice', 'BufferTime', 'Active') { $nr[$col] = [string]$r.$col }
			$nr['OrigScaleBrand'] = [string]$r.ScaleBrand
			$nr['OrigScaleModel'] = [string]$r.ScaleModel
			$nr['OrigIPNetwork'] = [string]$r.IPNetwork
			$nr['OrigIPDevice'] = [string]$r.IPDevice
			[void]$dt.Rows.Add($nr)
		}
	}

	# Reload the grid from the DB; $true when it worked
	$reloadFromDb = {
		param ([string]$What = 'refresh')
		try
		{
			& $fillGrid (& $runQuery $selectQuery)
			if ($grid) { $grid.Refresh() }
			return $true
		}
		catch
		{
			Write_Log "SQL error ($What): $($_.Exception.Message)" "red"
			return $false
		}
	}

	Write_Log "Loading table for edit..." "blue"
	if (-not (& $reloadFromDb 'load')) { return }
	if ($dt.Rows.Count -eq 0)
	{
		Write_Log "TBS_SCL_ver520 has no rows - nothing to edit." "yellow"
		return
	}

	# ----------------------------------------------------------------------------------------------
	# 5) WinForms UI - simple layout (no drag UI, no checkbox)
	# ----------------------------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Edit Placings - TBS_SCL_ver520"
	$form.StartPosition = 'CenterScreen'
	$form.Size = New-Object System.Drawing.Size(1100, 720)
	$form.MinimumSize = New-Object System.Drawing.Size(960, 600)
	$form.AutoScaleMode = 'Dpi'
	$form.TopMost = $true

	$grid = New-Object System.Windows.Forms.DataGridView
	$grid.Dock = [System.Windows.Forms.DockStyle]::Fill
	$grid.AutoGenerateColumns = $false
	$grid.AllowUserToAddRows = $false
	$grid.AllowUserToDeleteRows = $false
	$grid.ReadOnly = $false
	$grid.SelectionMode = 'FullRowSelect'
	$grid.MultiSelect = $false
	$grid.AutoSizeColumnsMode = 'Fill'
	$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
	$grid.AllowUserToOrderColumns = $true
	$grid.RowHeadersWidth = 30

	# CHANGE: Inline column creation (removed $AddCol helper).
	$col_ScaleCode = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleCode.HeaderText = 'ScaleCode'
	$col_ScaleCode.DataPropertyName = 'ScaleCode'
	$col_ScaleCode.ReadOnly = $false
	$col_ScaleCode.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleCode) | Out-Null

	$col_ScaleName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleName.HeaderText = 'ScaleName'
	$col_ScaleName.DataPropertyName = 'ScaleName'
	$col_ScaleName.ReadOnly = $false
	$col_ScaleName.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleName) | Out-Null

	$col_ScaleBrand = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleBrand.HeaderText = 'ScaleBrand'
	$col_ScaleBrand.DataPropertyName = 'ScaleBrand'
	$col_ScaleBrand.ReadOnly = $false
	$col_ScaleBrand.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleBrand) | Out-Null

	$col_ScaleModel = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_ScaleModel.HeaderText = 'ScaleModel'
	$col_ScaleModel.DataPropertyName = 'ScaleModel'
	$col_ScaleModel.ReadOnly = $false
	$col_ScaleModel.SortMode = 'NotSortable'
	$grid.Columns.Add($col_ScaleModel) | Out-Null

	$col_IPNetwork = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_IPNetwork.HeaderText = 'IPNetwork'
	$col_IPNetwork.DataPropertyName = 'IPNetwork'
	$col_IPNetwork.ReadOnly = $false
	$col_IPNetwork.SortMode = 'NotSortable'
	$grid.Columns.Add($col_IPNetwork) | Out-Null

	$col_IPDevice = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_IPDevice.HeaderText = 'IPDevice'
	$col_IPDevice.DataPropertyName = 'IPDevice'
	$col_IPDevice.ReadOnly = $false
	$col_IPDevice.SortMode = 'NotSortable'
	$grid.Columns.Add($col_IPDevice) | Out-Null

	$col_BufferTime = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_BufferTime.HeaderText = 'BufferTime'
	$col_BufferTime.DataPropertyName = 'BufferTime'
	$col_BufferTime.ReadOnly = $false
	$col_BufferTime.SortMode = 'NotSortable'
	$grid.Columns.Add($col_BufferTime) | Out-Null

	$col_Active = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_Active.HeaderText = 'Active'
	$col_Active.DataPropertyName = 'Active'
	$col_Active.ReadOnly = $false
	$col_Active.SortMode = 'NotSortable'
	$grid.Columns.Add($col_Active) | Out-Null

	# Hidden originals
	$col_OrigScaleBrand = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigScaleBrand.HeaderText = 'OrigScaleBrand'
	$col_OrigScaleBrand.DataPropertyName = 'OrigScaleBrand'
	$col_OrigScaleBrand.ReadOnly = $true
	$col_OrigScaleBrand.Visible = $false
	$col_OrigScaleBrand.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigScaleBrand) | Out-Null

	$col_OrigScaleModel = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigScaleModel.HeaderText = 'OrigScaleModel'
	$col_OrigScaleModel.DataPropertyName = 'OrigScaleModel'
	$col_OrigScaleModel.ReadOnly = $true
	$col_OrigScaleModel.Visible = $false
	$col_OrigScaleModel.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigScaleModel) | Out-Null

	$col_OrigIPNetwork = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigIPNetwork.HeaderText = 'OrigIPNetwork'
	$col_OrigIPNetwork.DataPropertyName = 'OrigIPNetwork'
	$col_OrigIPNetwork.ReadOnly = $true
	$col_OrigIPNetwork.Visible = $false
	$col_OrigIPNetwork.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigIPNetwork) | Out-Null

	$col_OrigIPDevice = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
	$col_OrigIPDevice.HeaderText = 'OrigIPDevice'
	$col_OrigIPDevice.DataPropertyName = 'OrigIPDevice'
	$col_OrigIPDevice.ReadOnly = $true
	$col_OrigIPDevice.Visible = $false
	$col_OrigIPDevice.SortMode = 'NotSortable'
	$grid.Columns.Add($col_OrigIPDevice) | Out-Null

	$grid.DataSource = $dt
	$form.Controls.Add($grid)

	# Bottom bar
	$bottom = New-Object System.Windows.Forms.Panel
	$bottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
	$bottom.Height = 86
	$form.Controls.Add($bottom)

	$lbl = New-Object System.Windows.Forms.Label
	$lbl.AutoSize = $true
	$lbl.Location = New-Object System.Drawing.Point(10, 10)
	if ($allowedPrefix)
	{
		# CHANGE: Reflect the enforced trailing period examples only.
		$lbl.Text = "Edit cells directly  |  Auto Sort renumbers and saves  |  Active: Y/N  |  IPNetwork: '$allowedPrefix.' (accepts $allowedPrefix, $allowedPrefix.0, $allowedPrefix.*, $allowedPrefix.x)"
	}
	else
	{
		# CHANGE: Note that a trailing period is still enforced if we can parse A.B.C.
		$lbl.Text = "Edit cells directly. Auto Sort applies standard order + numbering and saves immediately.  Active must be Y/N. IPNetwork will be normalized to 'A.B.C.' (trailing dot) if parsable."
	}
	$bottom.Controls.Add($lbl)

	$btnAuto = New-Object System.Windows.Forms.Button
	$btnAuto.Text = "Auto Sort (BIZERBA first)"
	$btnAuto.Size = New-Object System.Drawing.Size(220, 28)
	$btnAuto.Location = New-Object System.Drawing.Point(10, 44)
	$bottom.Controls.Add($btnAuto)

	$btnRefresh = New-Object System.Windows.Forms.Button
	$btnRefresh.Text = "Refresh"
	$btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
	$btnRefresh.Location = New-Object System.Drawing.Point(240, 44)
	$bottom.Controls.Add($btnRefresh)

	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = "Close"
	$btnClose.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnClose)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnCancel)

	$btnSave = New-Object System.Windows.Forms.Button
	$btnSave.Text = "Save"
	$btnSave.Size = New-Object System.Drawing.Size(100, 28)
	$bottom.Controls.Add($btnSave)

	# CHANGE: Inline initial placement (removed $placeButtons helper)
	$bw = [int]$bottom.ClientSize.Width
	$btnSave.Location = New-Object System.Drawing.Point(($bw - 110), 44)
	$btnCancel.Location = New-Object System.Drawing.Point(($bw - 220), 44)
	$btnClose.Location = New-Object System.Drawing.Point(($bw - 330), 44)

	# CHANGE: Inline resize handler computes placements directly
	$form.add_Resize({
			$bw2 = [int]$bottom.ClientSize.Width
			$btnSave.Location = New-Object System.Drawing.Point(($bw2 - 110), 44)
			$btnCancel.Location = New-Object System.Drawing.Point(($bw2 - 220), 44)
			$btnClose.Location = New-Object System.Drawing.Point(($bw2 - 330), 44)
		})

	# ----------------------------------------------------------------------------------------------
	# Save: validate + write exactly what is in the grid (shared by Save and Auto Sort)
	# ----------------------------------------------------------------------------------------------
	# ISNULL on the target keys: a scale with a NULL brand / model / IP part was never matched
	# (and silently not saved); the row count tells when a row still did not match.
	$sqlSaveTemplate = @'
DECLARE @Updated int = 0;
BEGIN TRY
  BEGIN TRAN;

  DECLARE @Placings TABLE(
      OrigScaleBrand nvarchar(100) NOT NULL,
      OrigScaleModel nvarchar(100) NOT NULL,
      OrigIPNetwork  nvarchar(100) NOT NULL,
      OrigIPDevice   nvarchar(100) NOT NULL,
      NewScaleBrand  nvarchar(100) NOT NULL,
      NewScaleModel  nvarchar(100) NOT NULL,
      NewIPNetwork   nvarchar(100) NOT NULL,
      NewIPDevice    nvarchar(100) NOT NULL,
      NewScaleCode   int           NOT NULL,
      NewBufferTime  nvarchar(20)  NOT NULL,
      NewScaleName   nvarchar(200) NOT NULL,
      NewActive      nvarchar(20)  NOT NULL
  );

  INSERT INTO @Placings (
      OrigScaleBrand, OrigScaleModel, OrigIPNetwork, OrigIPDevice,
      NewScaleBrand, NewScaleModel, NewIPNetwork, NewIPDevice,
      NewScaleCode, NewBufferTime, NewScaleName, NewActive
  )
  VALUES
  $(VALUES_BLOCK);

  UPDATE tgt
     SET tgt.ScaleBrand = src.NewScaleBrand,
         tgt.ScaleModel = src.NewScaleModel,
         tgt.IPNetwork  = src.NewIPNetwork,
         tgt.IPDevice   = src.NewIPDevice,
         tgt.ScaleCode  = src.NewScaleCode,
         tgt.BufferTime = src.NewBufferTime,
         tgt.ScaleName  = src.NewScaleName,
         tgt.Active     = src.NewActive
  FROM [TBS_SCL_ver520] AS tgt
  JOIN @Placings AS src
    ON  ISNULL(CAST(tgt.ScaleBrand AS nvarchar(100)), N'') = src.OrigScaleBrand
    AND ISNULL(CAST(tgt.ScaleModel AS nvarchar(100)), N'') = src.OrigScaleModel
    AND ISNULL(CAST(tgt.IPNetwork AS nvarchar(100)), N'')  = src.OrigIPNetwork
    AND ISNULL(CAST(tgt.IPDevice AS nvarchar(100)), N'')   = src.OrigIPDevice;
  SET @Updated = @@ROWCOUNT;

  COMMIT TRAN;
END TRY
BEGIN CATCH
  IF @@TRANCOUNT > 0 ROLLBACK TRAN;
  DECLARE @msg nvarchar(4000) = ERROR_MESSAGE();
  RAISERROR(@msg, 16, 1);
END CATCH;
SELECT @Updated;
'@

	$ipnPattern = '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$'

	# Validate every grid row (shows the first problem); normalizes Active and IPNetwork in place
	$validateGrid = {
		$seenCodes = @{ }
		$rowIdx = 0
		foreach ($row in $dt.Rows)
		{
			$rowIdx++
			$raw = $row['ScaleCode']
			$val = $null
			if (-not [int]::TryParse([string]$raw, [ref]$val))
			{
				[void][System.Windows.Forms.MessageBox]::Show("Invalid ScaleCode at row #${rowIdx}: '$raw' is not a number.", "Invalid Code", 'OK', 'Error')
				return $false
			}
			if ($seenCodes.ContainsKey($val))
			{
				[void][System.Windows.Forms.MessageBox]::Show("Duplicate ScaleCode '$val' detected (row #$rowIdx).", "Duplicate Code", 'OK', 'Error')
				return $false
			}
			$seenCodes[$val] = $true

			$actRaw = [string]$row['Active']
			$actUp = $actRaw.Trim().ToUpper()
			if ($actUp -ne 'Y' -and $actUp -ne 'N')
			{
				[void][System.Windows.Forms.MessageBox]::Show("Invalid Active value at row #${rowIdx}: '$actRaw'. Use Y or N.", "Invalid Active", 'OK', 'Error')
				return $false
			}
			$row['Active'] = $actUp

			# IPNetwork: A.B.C(.anything) normalized to "A.B.C." (always a trailing period);
			# with a detected NIC prefix it must be that /24
			$ipnRaw = ([string]$row['IPNetwork']).Trim()
			$m = [regex]::Match($ipnRaw, $ipnPattern)
			if (-not $m.Success)
			{
				$expect = if ($allowedPrefix) { "something within '$allowedPrefix.*'" } else { "a pattern like A.B.C or A.B.C.X" }
				[void][System.Windows.Forms.MessageBox]::Show("Invalid IPNetwork at row #${rowIdx}: '$ipnRaw'. Expected $expect.", "Invalid IPNetwork", 'OK', 'Error')
				return $false
			}
			$prefix = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)"
			if ($allowedPrefix -and $prefix -ne $allowedPrefix)
			{
				[void][System.Windows.Forms.MessageBox]::Show("IPNetwork at row #$rowIdx must be within '$allowedPrefix.*'. Got '$ipnRaw'.", "IPNetwork Out of Range", 'OK', 'Error')
				return $false
			}
			$row['IPNetwork'] = "$prefix."
		}
		return $true
	}

	# Validate, save, reload; $true when saved. $What names the action in the log.
	$saveGrid = {
		param ([string]$What)

		# Commit any in-progress grid edit
		if ($grid.IsCurrentCellInEditMode) { [void]$grid.EndEdit() }
		else { [void]$grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }

		if (-not (& $validateGrid)) { return $false }

		$sqlText = { param ($v) (([string]$v) -replace "[`r`n]", '') -replace "'", "''" }
		$values = New-Object System.Collections.Generic.List[string]
		foreach ($row in $dt.Rows)
		{
			$parts = @(
				'OrigScaleBrand', 'OrigScaleModel', 'OrigIPNetwork', 'OrigIPDevice', 'ScaleBrand', 'ScaleModel', 'IPNetwork', 'IPDevice' |
				ForEach-Object { "'" + (& $sqlText $row[$_]) + "'" }
			)
			$parts += [string][int]$row['ScaleCode']
			$parts += @('BufferTime', 'ScaleName', 'Active' | ForEach-Object { "'" + (& $sqlText $row[$_]) + "'" })
			$values.Add('(' + ($parts -join ',') + ')')
		}
		if ($values.Count -eq 0)
		{
			Write_Log "Nothing to save ($What)." "yellow"
			return $false
		}

		Write_Log "Saving to database ($What)..." "blue"
		try
		{
			$updated = [int](& $runScalar ($sqlSaveTemplate.Replace('$(VALUES_BLOCK)', [string]::Join(",`r`n  ", $values))))
		}
		catch
		{
			Write_Log "Save ($What) failed: $($_.Exception.Message)" "red"
			[void][System.Windows.Forms.MessageBox]::Show("Save failed. Check logs.", "Error", 'OK', 'Error')
			return $false
		}
		if ($updated -ne $values.Count) { Write_Log "Saved $updated of $($values.Count) row(s) - the others no longer match the table (changed elsewhere?). Refreshing." "yellow" }

		Write_Log "Refreshing view..." "blue"
		[void](& $reloadFromDb "refresh after $What")
		return $true
	}

	# ----------------------------------------------------------------------------------------------
	# 6) Button: Save (validate + write exactly what's in grid; no renumber)
	# ----------------------------------------------------------------------------------------------
	$btnSave.add_Click({
			if (& $saveGrid 'Save')
			{
				[void][System.Windows.Forms.MessageBox]::Show("Saved and refreshed.", "Success", 'OK', 'Information')
				Write_Log "Saved and refreshed." "green"
			}
		})

	# ----------------------------------------------------------------------------------------------
	# 7) Auto Sort: BIZERBA by numeric IP then ISHIDA, renumber, template names, immediate SAVE
	# ----------------------------------------------------------------------------------------------
	$btnAuto.add_Click({
			# Deep snapshots split by brand to avoid DataRow detach issues
			$snapB = @()
			$snapI = @()
			foreach ($row in $dt.Rows)
			{
				$o = [ordered]@{ }
				foreach ($col in $dt.Columns) { $o[$col.ColumnName] = $row[$col.ColumnName] }
				if ($o.ScaleBrand -eq 'BIZERBA') { $snapB += $o }
				elseif ($o.ScaleBrand -eq 'ISHIDA') { $snapI += $o }
			}

			# Sort by numeric IPDevice (non-numeric last), then by raw IPDevice
			$byDevice = @(
				@{ Expression = { $n = $null; if ([int]::TryParse([string]$_.IPDevice, [ref]$n)) { $n } else { [int]::MaxValue } } },
				@{ Expression = { $_.IPDevice } }
			)
			$snapB = @($snapB | Sort-Object $byDevice)
			$snapI = @($snapI | Sort-Object $byDevice)

			$dt.BeginLoadData()
			try
			{
				$dt.Rows.Clear()
				foreach ($s in @($snapB) + @($snapI)) { $nr = $dt.NewRow(); foreach ($k in $s.Keys) { $nr[$k] = [string]$s[$k] }; [void]$dt.Rows.Add($nr) }
			}
			finally { $dt.EndLoadData() }

			# Template naming + BufferTime for BIZERBA; Ishida WMAI name rule
			$firstB = $true
			foreach ($r in $dt.Rows)
			{
				if ($r['ScaleBrand'] -eq 'BIZERBA')
				{
					$r['ScaleName'] = "Scale $($r['IPDevice'])"
					if ($firstB) { $r['BufferTime'] = '1'; $firstB = $false }
					else { $r['BufferTime'] = '5' }
				}
			}
			$ishWmai = @($dt.Rows | Where-Object { $_['ScaleBrand'] -eq 'ISHIDA' -and $_['ScaleModel'] -eq 'WMAI' })
			if ($ishWmai.Count -gt 1) { foreach ($r in $ishWmai) { $r['ScaleName'] = "Ishida Wrapper $($r['IPDevice'])" } }
			elseif ($ishWmai.Count -eq 1) { $ishWmai[0]['ScaleName'] = 'Ishida Wrapper' }

			# Renumber: BIZERBA start at 10; ISHIDA start at (last BIZERBA)+10 or 20 if none
			$codeB = 10
			foreach ($r in $dt.Rows) { if ($r['ScaleBrand'] -eq 'BIZERBA') { $r['ScaleCode'] = [string]$codeB; $codeB = $codeB + 1 } }
			$ishidaStart = 20
			if ($codeB -gt 10) { $ishidaStart = ($codeB - 1) + 10 }
			$codeI = $ishidaStart
			foreach ($r in $dt.Rows) { if ($r['ScaleBrand'] -eq 'ISHIDA') { $r['ScaleCode'] = [string]$codeI; $codeI = $codeI + 1 } }

			$grid.Refresh()

			# Immediately SAVE (same validation rules as Save button)
			if (& $saveGrid 'Auto Sort')
			{
				[void][System.Windows.Forms.MessageBox]::Show("Auto Sort applied, saved, and refreshed.", "Success", 'OK', 'Information')
				Write_Log "Auto Sort applied, saved, and refreshed." "green"
			}
		})

	# ----------------------------------------------------------------------------------------------
	# 8) Refresh / Close / Cancel wiring
	# ----------------------------------------------------------------------------------------------
	$btnRefresh.add_Click({
			Write_Log "Refreshing view from database..." "blue"
			[void](& $reloadFromDb 'refresh')
		})

	$btnClose.add_Click({ $form.Close() })
	$btnCancel.add_Click({
			Write_Log "Cancel: reloading from database..." "yellow"
			$btnRefresh.PerformClick()
		})

	# ----------------------------------------------------------------------------------------------
	# 9) Show dialog
	# ----------------------------------------------------------------------------------------------
	[void]$form.ShowDialog()
	Write_Log "==================== Edit_TBS_SCL_ver520_Table Closed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Fix_Deploy_CHG
# ---------------------------------------------------------------------------------------------------
# Description:
#   Restores the deploy line in DEPLOY_CHG.sql: builds the @EXEC line from the FastDEPLOY
#   (preferred) or regular ScaleManagementApp.exe path, strips the existing deploy lines
#   (matched via the configured regex), then appends the new line. Also deletes the
#   Update_Scales_Specials_Task_Minutes scheduled task if it exists.
#
# Parameters:
#   - OfficePath  [string]: folder holding DEPLOY_CHG.sql (default: $script:OfficePath)
#   - Quiet       [switch]: no start/end banners (when called from another tool)
#
# Returns:
#   $true when the line was restored, otherwise $false.
#
# Notes:
#   - The one implementation: Update_Scales_Specials_Interactive (Restore button) and
#     Troubleshoot_ScaleCommApp call this instead of carrying their own copies.
#
# Variables:
#   - $script:ScaleManagementApp*PathCandidates: FastDEPLOY / regular exe lookup
#   - $script:ScaleCommDeployExecRegex: regex matching the @EXEC lines to strip
# ===================================================================================================

function Fix_Deploy_CHG
{
	param (
		[string]$OfficePath = $script:OfficePath,
		[switch]$Quiet
	)

	if (-not $Quiet) { Write_Log "`r`n==================== Starting Fix_Deploy_CHG Function ====================`r`n" "blue" }
	$endBanner = { if (-not $Quiet) { Write_Log "`r`n==================== Fix_Deploy_CHG Function Completed ====================" "blue" } }

	$deployChgFile = Join-Path $OfficePath "DEPLOY_CHG.sql"

	if (-not (Test-Path $deployChgFile))
	{
		Write_Log "DEPLOY_CHG.sql not found at $deployChgFile" "red"
		& $endBanner
		return $false
	}

	$exeLine = ""
	$fastDeployPath = $script:ScaleManagementAppFastDeployPathCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
	$regularDeployPath = $script:ScaleManagementAppPathCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
	if (-not $fastDeployPath) { $fastDeployPath = $script:ScaleManagementAppFastDeployPath }
	if (-not $regularDeployPath) { $regularDeployPath = $script:ScaleManagementAppExePath }
	if (Test-Path -LiteralPath $fastDeployPath)
	{
		$exeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='$fastDeployPath');"
	}
	elseif (Test-Path -LiteralPath $regularDeployPath)
	{
		$exeLine = "/* Deploy price changes to the scales */`r`n@EXEC(RUN='$regularDeployPath');"
	}
	else
	{
		Write_Log "Neither FastDEPLOY nor regular ScaleManagementApp.exe found in configured ScaleComm roots." "red"
		& $endBanner
		return $false
	}

	# --- Restore DEPLOY_CHG.sql ---
	try
	{
		$deployExecRegex = $script:ScaleCommDeployExecRegex
		$content = Get-Content $deployChgFile -Raw
		$newContent = [System.Collections.Generic.List[string]]@(
			($content -split "`r?`n") | Where-Object {
				$_ -notmatch '^\s*/\* Deploy price changes to the scales \*/' -and
				$_ -notmatch $deployExecRegex
			}
		)
		while ($newContent.Count -gt 0 -and ($newContent[-1] -match '^\s*$'))
		{
			$null = $newContent.RemoveAt($newContent.Count - 1)
		}
		$newContent += ""
		$newContent += $exeLine
		$newContent -join "`r`n" | Set-Content -Path $deployChgFile -Encoding Default -ErrorAction Stop
		Write_Log "Restored line to DEPLOY_CHG.sql: $exeLine" "green"
	}
	catch
	{
		Write_Log "Failed to restore line to DEPLOY_CHG.sql: $_" "red"
		& $endBanner
		return $false
	}

	# --- Delete the minutes task if it exists ---
	$minutesTaskName = "Update_Scales_Specials_Task_Minutes"
	$taskExists = schtasks /Query /TN "$minutesTaskName" 2>&1 | Select-String -Quiet -Pattern "$minutesTaskName"
	if ($taskExists)
	{
		$deleteOut = schtasks /Delete /TN "$minutesTaskName" /F 2>&1
		if ($LASTEXITCODE -eq 0)
		{
			Write_Log "Deleted scheduled task '$minutesTaskName' after DEPLOY_CHG.sql restore." "yellow"
		}
		else
		{
			Write_Log "Attempted to delete '$minutesTaskName' after restore but schtasks said: $deleteOut" "yellow"
		}
	}

	& $endBanner
	return $true
}

# ===================================================================================================
#               SECTION: Remote lane core (ported verbatim from the store maintenance script)
# ---------------------------------------------------------------------------------------------------
#   Source: Desktop\script.ps1 (Alex_C.T). Business logic UNCHANGED. Lane discovery (Retrieve_Nodes),
#   the lane picker (Show_Node_Selection_Form), the mailslot sender and the six remote lane tools.
#   They log through the Write_Log shim and read the paths set by Initialize_RemoteContext.
# ===================================================================================================

# ===================================================================================================
#                              FUNCTION: Ensure_MailslotSender_Type
# ---------------------------------------------------------------------------------------------------
# Description:
#   Loads the C# MailslotSender type (kernel32 CreateFile/WriteFile) used for direct Windows
#   mailslot messaging. Idempotent: a PSTypeName guard makes repeat calls no-ops, and it is a
#   function (not a bare top-level block) so background jobs can load it too.
#
# Notes:
#   - Called once at script top; any job that receives its body can call it as well.
#   - Exposes [MailslotSender]::SendMailslotCommand(mailslotName, command).
# ===================================================================================================

function Ensure_MailslotSender_Type
{
	if (-not ([System.Management.Automation.PSTypeName]'MailslotSender').Type)
	{
		Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class MailslotSender {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile
    );

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool WriteFile(
        IntPtr hFile,
        byte[] lpBuffer,
        uint nNumberOfBytesToWrite,
        out uint lpNumberOfBytesWritten,
        IntPtr lpOverlapped
    );
        
    [DllImport("kernel32.dll")]
    public static extern bool CloseHandle(IntPtr hObject);

    public static bool SendMailslotCommand(string mailslotName, string command) {
        const uint GENERIC_WRITE = 0x40000000;
        const uint FILE_SHARE_READ = 0x00000001;
        const uint OPEN_EXISTING = 3;

        IntPtr hFile = CreateFile(mailslotName, GENERIC_WRITE, FILE_SHARE_READ, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
        if (hFile == new IntPtr(-1)) {
            return false;
        }

        byte[] data = Encoding.ASCII.GetBytes(command);
        uint bytesWritten;
        bool success = WriteFile(hFile, data, (uint)data.Length, out bytesWritten, IntPtr.Zero);
        CloseHandle(hFile);
        return success;
    }
}
"@
	}
}

# ===================================================================================================
#                              FUNCTION: Delete_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deletes files/folders under a directory (recursively), optionally only those matching
#   name patterns and never those matching exclusion patterns. Runs synchronously (with log
#   lines) or as a background job so the caller is not held up.
#
# Parameters:
#   - Path           [string]: directory to delete from (required)
#   - SpecifiedFiles [string[]]: name patterns to delete (wildcards); all items when omitted
#   - Exclusions     [string[]]: name patterns to keep (wildcards)
#   - AsJob          [switch]: run the deletion as a background job
#
# Returns:
#   Number of items deleted (synchronous), or the Start-Job job (AsJob).
#
# Notes:
#   - Patterns match item NAMES at any depth (Get-ChildItem -Include/-Exclude): an excluded
#     folder name does not protect the files inside it.
# ===================================================================================================

function Delete_Files
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory, Position = 0, HelpMessage = "The directory path where files and folders will be deleted.")]
		[ValidateNotNullOrEmpty()]
		[string]$Path,
		[Parameter(HelpMessage = "Specific file/folder patterns to delete. Wildcards supported.")]
		[string[]]$SpecifiedFiles,
		[Parameter(HelpMessage = "Patterns to exclude from deletion. Wildcards supported.")]
		[string[]]$Exclusions,
		[Parameter(HelpMessage = "Run the deletion as a background job.")]
		[switch]$AsJob
	)

	# The deletion itself - shared by the job and the synchronous path.
	# Returns -1 when the path does not exist, else the number of items deleted.
	$deleteBlock = {
		param ($Path, $SpecifiedFiles, $Exclusions)

		$rp = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
		if (-not $rp) { return -1 }

		$gciParams = @{ Path = (Join-Path $rp.Path '*'); Recurse = $true; Force = $true; ErrorAction = 'SilentlyContinue' }
		if ($SpecifiedFiles -and $SpecifiedFiles.Count -gt 0) { $gciParams['Include'] = $SpecifiedFiles }
		if ($Exclusions -and $Exclusions.Count -gt 0) { $gciParams['Exclude'] = $Exclusions }

		$deletedCount = 0
		foreach ($it in @(Get-ChildItem @gciParams))
		{
			try
			{
				Remove-Item -LiteralPath $it.FullName -Force -Recurse -ErrorAction Stop
				$deletedCount++
			}
			catch { }
		}
		return $deletedCount
	}

	if ($AsJob)
	{
		$jobName = "DeleteFilesJob_{0}" -f (($Path -replace '[^A-Za-z0-9_-]', '_').Trim('_'))
		return Start-Job -Name $jobName -ScriptBlock $deleteBlock -ArgumentList $Path, $SpecifiedFiles, $Exclusions
	}

	try
	{
		$deletedCount = & $deleteBlock $Path $SpecifiedFiles $Exclusions
	}
	catch
	{
		Write_Log "An error occurred during deletion. $_" "red"
		return 0
	}

	if ($deletedCount -lt 0)
	{
		Write_Log "The specified path '$Path' does not exist." "red"
		return 0
	}
	if ($deletedCount -eq 0) { Write_Log "No items matched in '$Path'." "yellow" }
	Write_Log "Total items deleted: $deletedCount" "blue"
	return $deletedCount
}

# ===================================================================================================
#                              FUNCTION: Write_Sms_File
# ---------------------------------------------------------------------------------------------------
# Description:
#   Writes a file SMS will pick up (.sqi / .sql / .sqm drops in Office or a lane's XF folder) the
#   way SMS expects it: Windows-1252 (ANSI), CRLF line endings, a final CRLF, no BOM, and plain
#   attributes - no archive bit (SMS treats an archived file as already processed). A read-only
#   or hidden file of the same name is unlocked first (the write used to fail on those).
#
# Parameters:
#   - Path    [string]: file to write (required)
#   - Content [string]: file text, any line endings (required, may be empty)
#
# Returns:
#   Nothing. Throws when the file cannot be written - callers log their own message.
# ===================================================================================================

function Write_Sms_File
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Path,
		[Parameter(Mandatory = $true)]
		[AllowEmptyString()]
		[string]$Content
	)

	$text = $Content -replace "`r?`n", "`r`n"
	if ($text -and -not $text.EndsWith("`r`n")) { $text += "`r`n" }
	if ([System.IO.File]::Exists($Path)) { [System.IO.File]::SetAttributes($Path, [System.IO.FileAttributes]::Normal) }
	[System.IO.File]::WriteAllText($Path, $text, [System.Text.Encoding]::GetEncoding(1252))
	[System.IO.File]::SetAttributes($Path, [System.IO.FileAttributes]::Normal)
}

# ===================================================================================================
#                              FUNCTION: Get_All_Lanes_Database_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Collects per-lane SQL details (machine name, DB server, DB name) for every lane of the
#   store. Uses in-memory results first, then the per-function cache file, and only queries
#   the lanes when both are missing or -RefreshCache is set. With -LaneNumber it returns
#   that lane's info object.
#
# Parameters:
#   - LaneNumber   [string]: return only this lane's data (3-digit, padded)
#   - RefreshCache [switch]: force re-query and cache rewrite
#
# Variables:
#   - $script:FunctionResults['LaneDatabaseInfo']: per-lane info bucket (output)
#   - $script:FunctionResults['LaneNumToMachineName']: lane -> machine map (input)
#
# Returns:
#   The requested lane's info object; nothing for a full-store refresh.
#
# Writes:
#   - Setup_Files\Get_All_Lanes_Database_Info.txt (JSON cache, ordered by lane number)
# ===================================================================================================

function Get_All_Lanes_Database_Info
{
	param (
		[string]$LaneNumber,
		[switch]$RefreshCache
	)

	# Ensure bucket exists (unchanged behavior)
	if (-not $script:FunctionResults.ContainsKey('LaneDatabaseInfo'))
	{
		$script:FunctionResults['LaneDatabaseInfo'] = @{ }
	}
	$LaneDatabaseInfo = $script:FunctionResults['LaneDatabaseInfo']

	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName) { return $null }

	# Guard against null OfficePath causing Join-Path errors (fast fail, no guessing)
	if ([string]::IsNullOrWhiteSpace($OfficePath)) { return $null }

	# ------------------------------------------------------------------------------------------------
	# Per-function cache file (only run queries if memory/file data is missing or -RefreshCache is used)
	# ------------------------------------------------------------------------------------------------
	$CachePath = Join-Path $script:SetupFiles 'Get_All_Lanes_Database_Info.txt'

	# Normalize LaneNumber if provided (keep it simple + safe)
	$RequestedLane = $null
	if ($LaneNumber)
	{
		$digits = ($LaneNumber -replace '[^\d]', '')
		if ($digits) { $RequestedLane = $digits.PadLeft(3, '0') }
	}

	if (-not $RefreshCache)
	{
		# Reuse in-memory data first
		if ($RequestedLane -and $LaneDatabaseInfo.ContainsKey($RequestedLane))
		{
			return $LaneDatabaseInfo[$RequestedLane]
		}
		if ((-not $RequestedLane) -and $LaneDatabaseInfo.Count -gt 0)
		{
			return
		}

		# Reuse file cache if present
		if ($CachePath -and (Test-Path $CachePath))
		{
			try
			{
				$raw = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
				if (-not [string]::IsNullOrWhiteSpace($raw))
				{
					$cacheObj = $raw | ConvertFrom-Json -ErrorAction Stop
					if ($cacheObj -and $cacheObj.PSObject.Properties.Name -contains 'LaneDatabaseInfo' -and $cacheObj.LaneDatabaseInfo)
					{
						foreach ($p in $cacheObj.LaneDatabaseInfo.PSObject.Properties)
						{
							if (-not $LaneDatabaseInfo.ContainsKey($p.Name))
							{
								# Same shape as a fresh lookup (hashtable): callers index it as
								# $info['TcpConnStr'], which returns nothing on a PSCustomObject
								$laneHash = @{ }
								if ($p.Value -and $p.Value.PSObject) { foreach ($q in $p.Value.PSObject.Properties) { $laneHash[$q.Name] = $q.Value } }
								# Caches written before v2.7.0 lack TrustServerCertificate
								foreach ($k in @('ConnectionString', 'NamedPipesConnStr', 'TcpConnStr'))
								{
									if ($laneHash[$k] -and $laneHash[$k] -notmatch '(?i)TrustServerCertificate') { $laneHash[$k] = ([string]$laneHash[$k]).TrimEnd(';') + ';TrustServerCertificate=True;' }
								}
								$LaneDatabaseInfo[$p.Name] = $laneHash
							}
						}
					}
				}
			}
			catch { }

			if ($RequestedLane)
			{
				if ($LaneDatabaseInfo.ContainsKey($RequestedLane)) { return $LaneDatabaseInfo[$RequestedLane] }
			}
			elseif ($LaneDatabaseInfo.Count -gt 0)
			{
				return
			}
		}
	}

	# ------------------------------------------------------------------------------------------------
	# Memory/file cache is missing/unreadable, or caller requested -RefreshCache => query lanes ONCE, then write file
	# ------------------------------------------------------------------------------------------------
	$lanesToProcess = @()

	# Always build a full cache when missing
	$lanesToProcess += $LaneNumToMachineName.Keys

	# Ensure requested lane is included even if it would normally be skipped in "full mode"
	if ($RequestedLane) { $lanesToProcess = @($RequestedLane) + $lanesToProcess }

	# De-dupe + keep only numeric 3-digit lanes
	$lanesToProcess = $lanesToProcess | ForEach-Object {
		$d = ($_.ToString() -replace '[^\d]', '')
		if ($d) { $d.PadLeft(3, '0') }
	} | Where-Object { $_ } | Select-Object -Unique

	$changed = $false

	# Lanes to read: all of them with -RefreshCache (an entry is kept when its lane cannot be
	# read), otherwise only the ones not in memory yet. 8xx (scales) / 9xx (server, backoffices)
	# are not lanes - unless explicitly requested.
	$work = @()
	foreach ($laneNumber in $lanesToProcess)
	{
		if (-not $RefreshCache -and $LaneDatabaseInfo.ContainsKey($laneNumber)) { continue }
		if ($laneNumber -ne $RequestedLane -and $laneNumber -match '^(8|9)') { continue }
		$machineName = $LaneNumToMachineName[$laneNumber]
		if (-not $machineName) { continue }
		$work += [pscustomobject]@{ Lane = $laneNumber; Machine = [string]$machineName; Local = (Join-Path $OfficePath "INI\STARTUP.$laneNumber") }
	}

	# Lanes without a local STARTUP.<lane> copy are read over \\lane\storeman: probe port 445 of
	# all of them at once (1.5 s in total, not 1.5 s per switched-off lane)
	$reachable = @{ }
	$probes = @()
	foreach ($w in $work)
	{
		if (Test-Path -LiteralPath $w.Local) { continue }
		$client = New-Object System.Net.Sockets.TcpClient
		try { $probes += [pscustomobject]@{ Lane = $w.Lane; Client = $client; Task = $client.ConnectAsync($w.Machine, 445) } }
		catch { $client.Dispose() }
	}
	$deadline = (Get-Date).AddMilliseconds(1500)
	foreach ($p in $probes)
	{
		$ok = $false
		try
		{
			$left = [int][Math]::Max(0, ($deadline - (Get-Date)).TotalMilliseconds)
			$ok = ($p.Task.Wait($left) -and $p.Client.Connected)
		}
		catch { $ok = $false }
		$reachable[$p.Lane] = $ok
		try { $p.Client.Dispose() }
		catch { }
	}

	foreach ($w in $work)
	{
		$laneNumber = $w.Lane
		$machineName = $w.Machine

		# ---------- 1. LOCAL Startup file (INI\STARTUP.###), 2. REMOTE \\lane\storeman\Startup.ini ----------
		$content = $null
		$source = $null
		if (Test-Path -LiteralPath $w.Local)
		{
			try
			{
				$content = Get-Content -LiteralPath $w.Local -ErrorAction Stop
				$source = "local"
			}
			catch { $content = $null }
		}
		if (-not $content -and $reachable[$laneNumber])
		{
			try
			{
				$content = Get-Content -LiteralPath "\\$machineName\storeman\Startup.ini" -ErrorAction Stop
				$source = "remote"
			}
			catch { $content = $null }
		}
		if (-not $content) { continue }

		# First line only: a second DBNAME= line would turn the value into an array
		$dbNameLine = $content | Where-Object { $_ -match '^DBNAME=' } | Select-Object -First 1
		$dbServerLine = $content | Where-Object { $_ -match '^DBSERVER=' } | Select-Object -First 1

		if ($dbNameLine)
		{
			$dbName = ($dbNameLine -replace '^DBNAME=', '').Trim()
		}
		else
		{
			continue
		}

		if ($dbServerLine)
		{
			$dbServerRaw = ($dbServerLine -replace '^DBSERVER=', '').Trim()
		}
		else
		{
			$dbServerRaw = ""
		}

		# -------------- PATCH: Remove trailing backslash unless a named instance is present --------------
		if ($dbServerRaw -match '^[^\\]+\\$')
		{
			$dbServerRaw = $dbServerRaw.TrimEnd('\')
		}
		# PATCH: If (LOCAL) or localhost (any case), replace with actual machine name
		if ($dbServerRaw -match '^(?i)\(LOCAL\)$' -or $dbServerRaw -match '^(?i)localhost$' -or $dbServerRaw -eq "")
		{
			$dbServer = $machineName
		}
		else
		{
			$dbServer = $dbServerRaw
		}
		# END PATCH --------------------------------------------------------------------------------------

		# Parse instance name for Named Pipes/TCP logic
		$serverName = $dbServer
		$instanceName = $null
		if ($dbServer -match '\\')
		{
			$parts = $dbServer -split '\\'
			$serverName = $parts[0]
			$instanceName = $parts[1]
		}
		elseif ($dbServer -match ',')
		{
			$serverName = $dbServer
			$instanceName = $null
		}
		else
		{
			$serverName = $dbServer
			$instanceName = $null
		}

		# Build connection strings
		if ($instanceName -and $instanceName.ToUpper() -ne "MSSQLSERVER")
		{
			$namedPipes = "np:\\$serverName\pipe\MSSQL`$$instanceName\sql\query"
			$tcpServer = "$serverName\$instanceName"
		}
		else
		{
			$namedPipes = "np:\\$serverName\pipe\sql\query"
			$tcpServer = $serverName
		}

		# TrustServerCertificate: SqlServer module 22+ encrypts by default and the lanes use
		# self-signed certificates, so without it every Invoke-Sqlcmd to a lane is refused
		$tcpConnStr = "Server=$tcpServer;Database=$dbName;Integrated Security=True;TrustServerCertificate=True;"
		$namedPipesConnStr = "Server=$namedPipes;Database=$dbName;Integrated Security=True;TrustServerCertificate=True;"
		$simpleConnStr = "Server=$dbServer;Database=$dbName;Integrated Security=True;TrustServerCertificate=True;"

		$laneInfo = @{
			'MachineName'	    = $machineName
			'DBName'		    = $dbName
			'DBServer'		    = $dbServer
			'ServerName'	    = $serverName
			'InstanceName'	    = $instanceName
			'NamedPipes'	    = $namedPipes
			'TcpServer'		    = $tcpServer
			'ConnectionString'  = $simpleConnStr
			'NamedPipesConnStr' = $namedPipesConnStr
			'TcpConnStr'	    = $tcpConnStr
			'Source'		    = $source
		}

		$LaneDatabaseInfo[$laneNumber] = $laneInfo
		$changed = $true
	}

	# ------------------------------------------------------------------------------------------------
	# Write cache ONCE (only if we found anything new)
	#   IMPORTANT: Write LaneDatabaseInfo ordered by lane number for clean file output.
	# ------------------------------------------------------------------------------------------------
	if ($CachePath -and $changed)
	{
		try
		{
			# Build an ORDERED view for file output only (keep in-memory as normal hashtable for .ContainsKey())
			$orderedLaneDb = [ordered]@{ }
			$sortedLaneKeys = @($LaneDatabaseInfo.Keys) | Sort-Object { [int]($_) }
			foreach ($k in $sortedLaneKeys)
			{
				$orderedLaneDb[$k] = $LaneDatabaseInfo[$k]
			}

			$cacheOut = [pscustomobject]@{
				LastUpdated = (Get-Date).ToString('o')
				LaneDatabaseInfo = $orderedLaneDb
			}

			($cacheOut | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
		}
		catch { }
	}

	# Return behavior (same intent as original):
	if ($RequestedLane)
	{
		if ($LaneDatabaseInfo.ContainsKey($RequestedLane)) { return $LaneDatabaseInfo[$RequestedLane] }
		return $null
	}
}

# ===================================================================================================
#                              FUNCTION: Retrieve_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Finds the store's lanes, scales, server and backoffices and builds every lookup map the
#   rest of the script uses (lane number -> machine, scale code -> IP info, ...). Sources, in
#   order: the store database (TER_TAB + TBS_SCL_ver520), else Office\Load\Ter_Load.sql; TER
#   scales are enriched from XchScale\XchNodes.ini. Results are cached per store, the GUI
#   count labels are updated, and the nodes object is returned.
#
# Parameters:
#   - StoreNumber  [string]: store to count (required)
#   - RefreshCache [switch]: ignore the in-memory / file cache and rediscover
#
# Variables:
#   - $script:FunctionResults: receives Nodes, every map listed in $resultNames below and the
#     NumberOfLanes/Servers/Backoffices/Scales counts
#   - Reads $OfficePath, $LoadPath, $script:BasePath, FunctionResults DBSERVER / DBNAME /
#     ConnectionString / SqlModuleName / StoremanPath
#
# Writes:
#   - Setup_Files\Retrieve_Nodes.txt (JSON cache; CacheSchemaVersion must match to be used)
#
# Returns:
#   PSCustomObject with all counts and the lane/scale/backoffice maps.
#
# Notes:
#   - Node rules (TER_TAB / Ter_Load rows): 0xx (or a \\machine\ path) = lane, 8xx with an
#     XchScale.exe path = TER scale, 901 '@machine' = server, 902-998 '@machine' = backoffice.
#   - TBS scales keep their original ScaleCode; a TER scale with the same IP as a TBS scale
#     is dropped (the TBS row wins and inherits the TER path/port details).
#   - Run as a background job at startup (warmup): keep it free of GUI-only dependencies
#     (the label updates are skipped when the labels do not exist).
# ===================================================================================================

function Retrieve_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[switch]$RefreshCache
	)

	# Bump when the shape of the cached results changes (10: TBS scales keep their original
	# ScaleCode and are never normalised to 80X; 11: TBS FullIP joins network + device with a dot)
	$CacheSchemaVersion = 11
	$CachePath = Join-Path $script:SetupFiles 'Retrieve_Nodes.txt'

	# ------------------------------------------------------------------------------------------------
	# What this function publishes. FunctionResults keys = local variable names.
	# ------------------------------------------------------------------------------------------------
	# Nodes object properties, in order
	$nodeNames = @(
		'NumberOfLanes', 'NumberOfServers', 'NumberOfBackoffices', 'NumberOfScales',
		'LaneMachineNames', 'LaneNumToMachineName', 'LaneMachineLabels', 'LaneMachinePath', 'LaneMachineToServerPath',
		'ScaleCodes', 'ScaleLabels', 'ScaleExePaths', 'ScaleCodeToIPInfo', 'ScaleNameToCode', 'ScaleCodeToPath', 'ScaleNameToPath',
		'ScaleCodeToName', 'ScaleCodeToSection', 'ScaleCodeToBrand', 'ScaleCodeToModel', 'ScaleCodeToLocation',
		'ScaleCodeToIPNetwork', 'ScaleCodeToIPDevice', 'ScaleCodeToFullIP', 'ScaleCodeToPortRaw', 'ScaleCodeToTcpPort',
		'ScaleCodeToBaseMode', 'ScaleCodeToType', 'ScaleCodeToBaud', 'ScaleCodeToSub', 'ScaleCodeToPlumDir',
		'BackofficeNumToMachineName', 'BackofficeNumToLabel', 'BackofficeNumToPath',
		'ServerMachineName', 'ServerLabel', 'ServerPath', 'WindowsScales'
	)
	# Published to FunctionResults (the node properties plus the reverse maps)
	$resultNames = $nodeNames + @('MachineNameToLaneNum', 'MachineNameToBackofficeNum', 'MachineNameToBackofficePath')
	# Hashtable maps (cache: JSON object <-> hashtable)
	$mapNames = @($resultNames | Where-Object { $_ -notmatch '^(NumberOf|ServerMachineName$|ServerLabel$|ServerPath$|LaneMachineNames$|ScaleCodes$)' })
	# Scale fields that also get their own ScaleCodeTo<Field> map (only when set)
	$scaleFieldMaps = @(
		@('Section', 'ScaleCodeToSection'), @('ScaleBrand', 'ScaleCodeToBrand'), @('ScaleModel', 'ScaleCodeToModel'),
		@('ScaleLocation', 'ScaleCodeToLocation'), @('IPNetwork', 'ScaleCodeToIPNetwork'), @('IPDevice', 'ScaleCodeToIPDevice'),
		@('FullIP', 'ScaleCodeToFullIP'), @('PortRaw', 'ScaleCodeToPortRaw'), @('TcpPort', 'ScaleCodeToTcpPort'),
		@('BaseMode', 'ScaleCodeToBaseMode'), @('Type', 'ScaleCodeToType'), @('Baud', 'ScaleCodeToBaud'),
		@('Sub', 'ScaleCodeToSub'), @('PlumDir', 'ScaleCodeToPlumDir')
	)

	# Count labels (GUI only; absent in background jobs)
	$updateLabels = {
		param ($Servers, $Backoffices, $Lanes, $Scales)
		if ($null -ne $NodesBackoffices) { $NodesBackoffices.Text = "Number of Backoffices: $Backoffices" }
		if ($null -ne $NodesStore) { $NodesStore.Text = "Number of Lanes: $Lanes" }
		if ($null -ne $scalesLabel) { $scalesLabel.Text = "Number of Scales: $Scales" }
		if ($null -ne $script:MainForm) { $script:MainForm.Refresh() }
	}

	# Builds $Nodes from the local variables and publishes everything (dot-sourced: runs in
	# this function's scope)
	$publish = {
		$nodeProps = [ordered]@{ }
		foreach ($n in $nodeNames) { $nodeProps[$n] = Get-Variable -Name $n -ValueOnly }
		$Nodes = [PSCustomObject]$nodeProps
		foreach ($n in $resultNames) { $script:FunctionResults[$n] = Get-Variable -Name $n -ValueOnly }
		$script:FunctionResults['Nodes'] = $Nodes
		$script:FunctionResults['XchNodesIniPathsUsed'] = @($XchNodesIniPathsUsed | Select-Object -Unique)
		& $updateLabels $NumberOfServers $NumberOfBackoffices $NumberOfLanes $NumberOfScales
	}

	# Store number as used for the cache match (3 digits, 4+ digit stores keep theirs)
	$normalizeStore = {
		param ([string]$Value)
		$digits = ([string]$Value -replace '[^\d]', '')
		if (-not $digits) { return $Value }
		try
		{
			if ($digits.Length -le 3) { return ([int]$digits).ToString('000') }
			return ([int]$digits).ToString('0000')
		}
		catch { return $Value }
	}
	$StoreNumberNorm = & $normalizeStore $StoreNumber

	# ------------------------------------------------------------------------------------------------
	# 0a. In-memory results for this store (this launch already discovered them)
	# ------------------------------------------------------------------------------------------------
	if (-not $RefreshCache)
	{
		try
		{
			$existingStoreNorm = & $normalizeStore ([string]$script:FunctionResults['StoreNumber'])
			if ($script:FunctionResults['Nodes'] -and ($existingStoreNorm -eq $StoreNumberNorm) -and $script:FunctionResults['LaneNumToMachineName'])
			{
				& $updateLabels $script:FunctionResults['NumberOfServers'] $script:FunctionResults['NumberOfBackoffices'] $script:FunctionResults['NumberOfLanes'] $script:FunctionResults['NumberOfScales']
				return $script:FunctionResults['Nodes']
			}
		}
		catch { }
	}

	# ------------------------------------------------------------------------------------------------
	# 0b. File cache for this store
	# ------------------------------------------------------------------------------------------------
	if ((-not $RefreshCache) -and (Test-Path -LiteralPath $CachePath))
	{
		try
		{
			$raw = Get-Content -LiteralPath $CachePath -Raw -ErrorAction Stop
			$cacheObj = $null
			if (-not [string]::IsNullOrWhiteSpace($raw)) { $cacheObj = $raw | ConvertFrom-Json -ErrorAction Stop }
			if ($cacheObj -and $cacheObj.Results -and $cacheObj.CacheSchemaVersion -and
				([int]$cacheObj.CacheSchemaVersion -eq $CacheSchemaVersion) -and
				$cacheObj.StoreNumber -and ($cacheObj.StoreNumber.ToString() -eq $StoreNumberNorm))
			{
				$r = $cacheObj.Results
				foreach ($n in $mapNames)
				{
					$h = @{ }
					$o = $r.$n
					if ($o -and $o.PSObject -and $o.PSObject.Properties) { foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = $p.Value } }
					Set-Variable -Name $n -Value $h
				}
				$NumberOfLanes = [int]$r.NumberOfLanes
				$NumberOfServers = [int]$r.NumberOfServers
				$NumberOfBackoffices = [int]$r.NumberOfBackoffices
				$NumberOfScales = [int]$r.NumberOfScales
				$LaneMachineNames = @($r.LaneMachineNames)
				$ScaleCodes = @($r.ScaleCodes)
				$ServerMachineName = $r.ServerMachineName
				$ServerLabel = $r.ServerLabel
				$ServerPath = $r.ServerPath
				$XchNodesIniPathsUsed = @()
				if ($r.PSObject.Properties.Name -contains 'XchNodesIniPathsUsed') { $XchNodesIniPathsUsed = @($r.XchNodesIniPathsUsed) }

				. $publish
				return $Nodes
			}
		}
		catch
		{
			# unreadable / old cache: discover again below
		}
	}

	# ====================================================================================
	# DISCOVERY - start empty
	# ====================================================================================
	foreach ($n in $mapNames) { Set-Variable -Name $n -Value @{ } }
	$LaneMachineNames = @()
	$ScaleCodes = @()
	$ServerMachineName = $null
	$ServerLabel = $null
	$ServerPath = $null
	$NumberOfLanes = 0
	$NumberOfServers = 0
	$NumberOfBackoffices = 0
	$NumberOfScales = 0

	# TER scales are grouped by their 80X terminal; TBS scales keep the ScaleCode stored in
	# TBS_SCL_ver520.
	$TerScaleBy80X = @{ }
	$TbsScaleRows = New-Object System.Collections.ArrayList

	# Classify one TER_TAB / Ter_Load row (dot-sourced: updates this function's variables)
	$addTerRow = {
		param ([string]$terminal, [string]$label, [string]$path, [string]$hostPath)

		if (($terminal -match '^0\d\d$' -or $path -match '^\\\\[^\\]+\\') -and $terminal -notmatch '^8' -and $terminal -notmatch '^9')
		{
			# Lane
			$machineName = if ($path -match '\\\\([^\\]+)\\') { $matches[1] } else { $terminal }
			$LaneMachineNames += $machineName
			if ($terminal -match '^0\d\d$') { $LaneNumToMachineName[$terminal] = $machineName }
			$LaneMachineLabels[$machineName] = $label
			$LaneMachinePath[$machineName] = $path
			$LaneMachineToServerPath[$machineName] = $hostPath
		}
		elseif ($terminal -match '^8\d\d$' -and $path -match '(?i)XchScale\\XchScale\.exe$')
		{
			# TER scale (one entry per 80X terminal; later rows only fill gaps)
			if (-not $TerScaleBy80X.ContainsKey($terminal))
			{
				$TerScaleBy80X[$terminal] = [PSCustomObject]@{
					Code          = $terminal
					ScaleCode     = $terminal
					BaseScaleCode = $terminal
					Label         = $label
					ScaleName     = $label
					Path          = $path
					HostPath      = $hostPath
					ScaleBrand    = $null
					Scalevendor   = $null
					ScaleModel    = $null
					ScaleLocation = $null
					IPNetwork     = $null
					IPDevice      = $null
					FullIP        = $null
					PortRaw       = $null
					TcpPort       = $null
					Section       = $null
					BaseMode      = $null
					Type          = $null
					Baud          = $null
					Sub           = $null
					PlumDir       = $null
					Active        = $null
					Source        = 'TER'
				}
			}
			else
			{
				$existingTer = $TerScaleBy80X[$terminal]
				if ([string]::IsNullOrWhiteSpace([string]$existingTer.Label) -and $label) { $existingTer.Label = $label }
				if ([string]::IsNullOrWhiteSpace([string]$existingTer.ScaleName) -and $label) { $existingTer.ScaleName = $label }
				if ([string]::IsNullOrWhiteSpace([string]$existingTer.Path) -and $path) { $existingTer.Path = $path }
				if ([string]::IsNullOrWhiteSpace([string]$existingTer.HostPath) -and $hostPath) { $existingTer.HostPath = $hostPath }
			}
		}
		elseif ($terminal -eq '901' -and $path -match '^@[^@]+$')
		{
			# Store server
			$ServerMachineName = $path -replace '^@', ''
			$ServerLabel = $label
			$ServerPath = $path
		}
		elseif ($terminal -match '^9(0[2-9]|[1-8]\d|9[0-8])$' -and $path -match '^@[^@]+$')
		{
			# Backoffice
			$BackofficeNumToMachineName[$terminal] = $path -replace '^@', ''
			$BackofficeNumToLabel[$terminal] = $label
			$BackofficeNumToPath[$terminal] = $path
		}
	}

	$TerLoadSqlPath = $null
	if (-not [string]::IsNullOrWhiteSpace($LoadPath)) { $TerLoadSqlPath = Join-Path $LoadPath 'Ter_Load.sql' }

	# ====================================================================================
	# XchNodes.ini candidates: <storeman>\XchScale (+ one or two folder levels below)
	# ====================================================================================
	$XchSearchRoots = @()
	foreach ($rootCandidate in @([string]$script:FunctionResults['StoremanPath'], [string]$script:BasePath))
	{
		if (-not [string]::IsNullOrWhiteSpace($rootCandidate)) { $XchSearchRoots += [System.IO.Path]::Combine($rootCandidate.Trim().TrimEnd('\'), 'XchScale') }
	}
	if (-not [string]::IsNullOrWhiteSpace($OfficePath))
	{
		try
		{
			$officeTrimmed = ([string]$OfficePath).Trim().TrimEnd('\')
			if ((Split-Path -Path $officeTrimmed -Leaf) -ieq 'Office')
			{
				$storemanRoot = Split-Path -Path $officeTrimmed -Parent
				if (-not [string]::IsNullOrWhiteSpace($storemanRoot)) { $XchSearchRoots += [System.IO.Path]::Combine($storemanRoot.TrimEnd('\'), 'XchScale') }
			}
			$XchSearchRoots += [System.IO.Path]::Combine($officeTrimmed, 'XchScale')
		}
		catch { }
	}

	$XchNodesIniCandidates = New-Object System.Collections.Generic.List[string]
	$seenIni = @{ }
	$addIni = {
		param ([string]$IniPath)
		if ([System.IO.File]::Exists($IniPath) -and -not $seenIni.ContainsKey($IniPath.Trim().ToLowerInvariant()))
		{
			$seenIni[$IniPath.Trim().ToLowerInvariant()] = $true
			$XchNodesIniCandidates.Add($IniPath)
		}
	}
	$seenRoots = @{ }
	foreach ($root in $XchSearchRoots)
	{
		$rootPath = ([string]$root).Trim().Trim('"')
		if ([string]::IsNullOrWhiteSpace($rootPath) -or $seenRoots.ContainsKey($rootPath.ToLowerInvariant())) { continue }
		$seenRoots[$rootPath.ToLowerInvariant()] = $true
		try
		{
			if (-not (Test-Path -LiteralPath $rootPath -PathType Container)) { continue }
			& $addIni ([System.IO.Path]::Combine($rootPath, 'XchNodes.ini'))

			# Bounded search (some stores have very large XchScale trees): children, then
			# grandchildren only if nothing was found yet
			$childDirs = @(Get-ChildItem -LiteralPath $rootPath -Directory -ErrorAction SilentlyContinue)
			foreach ($childDir in $childDirs) { & $addIni ([System.IO.Path]::Combine([string]$childDir.FullName, 'XchNodes.ini')) }
			if ($XchNodesIniCandidates.Count -eq 0)
			{
				foreach ($childDir in $childDirs)
				{
					foreach ($grandChildDir in @(Get-ChildItem -LiteralPath $childDir.FullName -Directory -ErrorAction SilentlyContinue))
					{
						& $addIni ([System.IO.Path]::Combine([string]$grandChildDir.FullName, 'XchNodes.ini'))
					}
				}
			}
		}
		catch { }
	}

	# ====================================================================================
	# 1. PRIMARY: STORE DATABASE (TER_TAB, TBS_SCL_ver520)
	# ====================================================================================
	$server = $script:FunctionResults['DBSERVER']
	$database = $script:FunctionResults['DBNAME']
	$ConnectionString = Normalize_Sql_ConnectionString -RawConnectionString $script:FunctionResults['ConnectionString'] -Server $server -Database $database -SqlUser $script:FunctionResults['SQL_UserID'] -SqlPwd $script:FunctionResults['SQL_UserPwd']
	$NodesFromDatabase = $false
	$SqlModule = $script:FunctionResults['SqlModuleName']

	if (-not $SqlModule -or $SqlModule -eq 'None')
	{
		Write_Log "SQL PowerShell module not installed - using the built-in SQL client for node info." "yellow"
	}

	# The lane query filters on this store (TER_TAB.F1056); without a known store number the
	# connection is a placeholder target and would only burn ~15 s on timeouts.
	$hasUsableStoreNumber = (([string]$StoreNumber -replace '[^\d]', '').Length -gt 0)
	if ($ConnectionString -and -not $hasUsableStoreNumber)
	{
		Write_Log "Store number unknown - skipping the store database lookup (no store to query against). Using local files instead." "yellow"
	}

	if ($ConnectionString -and $hasUsableStoreNumber)
	{
		try
		{
			if ($SqlModule -and $SqlModule -ne 'None') { Import-Module $SqlModule -ErrorAction SilentlyContinue | Out-Null }
		}
		catch { }
		$invokeSqlCmd = Get-Command Invoke-Sqlcmd -ErrorAction SilentlyContinue
		$supportsConnStr = $false
		$supportsConnTimeout = $false
		$startupQueryTimeoutSeconds = 10
		$startupConnectionTimeoutSeconds = 5
		$startupConnectionString = $ConnectionString
		if ($invokeSqlCmd)
		{
			$supportsConnStr = $invokeSqlCmd.Parameters.ContainsKey('ConnectionString')
			$supportsConnTimeout = $invokeSqlCmd.Parameters.ContainsKey('ConnectionTimeout')
		}
		elseif ($SqlModule -and $SqlModule -ne 'None')
		{
			Write_Log "SQL module '$SqlModule' not available on this machine - using the built-in SQL client." "yellow"
		}
		else
		{
			Write_Log "SQL module not available on this machine - using the built-in SQL client." "yellow"
		}

		# Invoke-Sqlcmd when available, else SqlClient (returns rows either way)
		$invokeStartupQuery = {
			param ([string]$Query, [int]$QueryTimeout = 10)

			if ($invokeSqlCmd)
			{
				if ($supportsConnStr) { return (& $invokeSqlCmd -ConnectionString $startupConnectionString -Query $Query -QueryTimeout $QueryTimeout -ErrorAction Stop) }
				if ($supportsConnTimeout) { return (& $invokeSqlCmd -ServerInstance $server -Database $database -Query $Query -QueryTimeout $QueryTimeout -ConnectionTimeout $startupConnectionTimeoutSeconds -ErrorAction Stop) }
				return (& $invokeSqlCmd -ServerInstance $server -Database $database -Query $Query -QueryTimeout $QueryTimeout -ErrorAction Stop)
			}

			$dt = New-Object System.Data.DataTable
			$cn = New-Object System.Data.SqlClient.SqlConnection $startupConnectionString
			try
			{
				$cn.Open()
				$cmd = $cn.CreateCommand()
				$cmd.CommandText = $Query
				$cmd.CommandTimeout = $QueryTimeout
				$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
				[void]$da.Fill($dt)
			}
			finally
			{
				try { if ($cn) { $cn.Close() } } catch { }
				try { if ($cn) { $cn.Dispose() } } catch { }
			}
			return $dt
		}

		# Short connect timeout for the startup queries
		if (-not [string]::IsNullOrWhiteSpace($startupConnectionString))
		{
			try
			{
				$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder $startupConnectionString
				$csb.PSBase.ConnectTimeout = $startupConnectionTimeoutSeconds
				$startupConnectionString = $csb.ConnectionString
			}
			catch
			{
				if ($startupConnectionString -notmatch '(?i)\bConnect\s*Timeout\s*=')
				{
					$startupConnectionString = $startupConnectionString.Trim().TrimEnd(';') + ";Connect Timeout=$startupConnectionTimeoutSeconds"
				}
			}
		}

		$NodesFromDatabase = $true
		$safeStoreNumber = ($StoreNumber -replace '[^\d]', '')
		if ($safeStoreNumber.Length -le 3) { $safeStoreNumber = ([int]$safeStoreNumber).ToString('000') }
		else { $safeStoreNumber = ([int]$safeStoreNumber).ToString('0000') }
		try
		{
			$terTabResult = & $invokeStartupQuery "SELECT F1057, F1058, F1125, F1169 FROM TER_TAB WHERE F1056 = '$safeStoreNumber'" $startupQueryTimeoutSeconds
			foreach ($row in $terTabResult)
			{
				. $addTerRow ([string]$row.F1057) ([string]$row.F1058) ([string]$row.F1125) ([string]$row.F1169)
			}

			$tbsSclScalesResult = $null
			try
			{
				$tbsSclScalesResult = & $invokeStartupQuery "SELECT ScaleCode, ScaleName, ScaleLocation, IPNetwork, IPDevice, Active, ScaleBrand, ScaleModel FROM TBS_SCL_ver520 WHERE Active = 'Y'" $startupQueryTimeoutSeconds
			}
			catch
			{
				# A store without the TBS scale table simply has no TBS scales
				if ($_.Exception.Message -notmatch "Invalid object name 'TBS_SCL_ver520'") { throw }
			}

			# Column value as string, $null when empty / DBNull
			$col = {
				param ($v)
				if ([string]::IsNullOrWhiteSpace([string]$v)) { return $null }
				return [string]$v
			}
			$tbsRowOrdinal = 0
			foreach ($row in @($tbsSclScalesResult))
			{
				if ($null -eq $row) { continue }
				$tbsRowOrdinal++
				$rawScaleCode = ([string]$row.ScaleCode).Trim()
				if ([string]::IsNullOrWhiteSpace($rawScaleCode)) { continue }

				$tbsScaleName = & $col $row.ScaleName
				$tbsIPNetwork = & $col $row.IPNetwork
				if ($tbsIPNetwork) { $tbsIPNetwork = $tbsIPNetwork.Trim() }
				$tbsIPDevice = & $col $row.IPDevice
				if ($tbsIPDevice) { $tbsIPDevice = $tbsIPDevice.Trim() }
				$tbsBrand = & $col $row.ScaleBrand
				$tbsLabel = if ($tbsScaleName) { $tbsScaleName } else { "Scale $rawScaleCode" }

				$null = $TbsScaleRows.Add([PSCustomObject]@{
						RawScaleCode  = $rawScaleCode
						TbsRowOrdinal = $tbsRowOrdinal
						ScaleCode     = $rawScaleCode
						BaseScaleCode = $rawScaleCode
						Label         = $tbsLabel
						ScaleName     = $tbsLabel
						Path          = $null
						HostPath      = $null
						ScaleBrand    = $tbsBrand
						Scalevendor   = $tbsBrand
						ScaleModel    = & $col $row.ScaleModel
						ScaleLocation = & $col $row.ScaleLocation
						IPNetwork     = $tbsIPNetwork
						IPDevice      = $tbsIPDevice
						# One dot between the parts: "192.168.1" + "50" must not become "192.168.150"
						# (Windows reads that as 192.168.0.150)
						FullIP        = if ($tbsIPNetwork -and $tbsIPDevice) { $tbsIPNetwork.TrimEnd('.') + '.' + $tbsIPDevice.TrimStart('.') }
						elseif ($tbsIPNetwork) { $tbsIPNetwork.TrimEnd('.') }
						elseif ($tbsIPDevice) { $tbsIPDevice }
						else { $null }
						PortRaw       = $null
						TcpPort       = $null
						Section       = $null
						BaseMode      = $null
						Type          = $null
						Baud          = $null
						Sub           = $null
						PlumDir       = $null
						Active        = & $col $row.Active
						Source        = 'TBS'
					})
			}
		}
		catch
		{
			Write_Log "Failed to retrieve counts from the database: $_" "yellow"
			$NodesFromDatabase = $false
		}
	}

	# ====================================================================================
	# 2. FALLBACK: Office\Load\Ter_Load.sql
	# ====================================================================================
	if ((-not $NodesFromDatabase) -and $TerLoadSqlPath -and (Test-Path -LiteralPath $TerLoadSqlPath))
	{
		Write_Log "Using Ter_Load.sql as backup for TER_TAB information." "yellow"
		$content = Get-Content -LiteralPath $TerLoadSqlPath -Raw
		# SMS load files put one row per line:  INSERT INTO Ter_Load VALUES<CRLF>(...),<CRLF>(...);
		# (the old single-line match only ever saw the first row = the first lane)
		$insert = [regex]::Match([string]$content, '(?is)INSERT\s+INTO\s+Ter_Load\s+VALUES\s*(.*?\))\s*;')
		if (-not $insert.Success) { $insert = [regex]::Match([string]$content, '(?is)INSERT\s+INTO\s+Ter_Load\s+VALUES\s*(.*)$') }
		if ($insert.Success)
		{
			foreach ($row in [regex]::Matches($insert.Groups[1].Value, "\((?<r>(?:'(?:[^']|'')*'|[^'()])*)\)"))
			{
				# Quoted values ('' = an escaped quote; commas and parentheses inside quotes are fine)
				$fields = @([regex]::Matches($row.Groups['r'].Value, "'((?:[^']|'')*)'") | ForEach-Object { $_.Groups[1].Value.Replace("''", "'") })
				if ($fields.Count -ge 5 -and ($fields[0] -eq '@STORE' -or (& $normalizeStore $fields[0]) -eq $StoreNumberNorm))
				{
					. $addTerRow ([string]$fields[1]) ([string]$fields[2]) ([string]$fields[3]) ([string]$fields[4])
				}
			}
		}
	}

	# ====================================================================================
	# 2.5 ENRICH TER 80X SCALES FROM XchNodes.ini (best section per scale wins)
	# ====================================================================================
	$XchNodesIniPathsUsed = @()
	if ($TerScaleBy80X.Count -gt 0 -and $XchNodesIniCandidates.Count -gt 0)
	{
		$AllowedScaleCodeSet = @{ }
		foreach ($allowedCode in @($TerScaleBy80X.Keys | Where-Object { $_ -match '^8\d\d$' })) { $AllowedScaleCodeSet[$allowedCode] = $true }

		# Section names are <store><terminal>, e.g. 001801 or 0001801
		$storeSectionPrefixes = @()
		$storeDigits = ($StoreNumber -replace '[^\d]', '')
		if ($storeDigits)
		{
			try { $storeSectionPrefixes += ([int]$storeDigits).ToString('000'); $storeSectionPrefixes += ([int]$storeDigits).ToString('0000') }
			catch { }
			$storeSectionPrefixes += $storeDigits
			$storeSectionPrefixes = @($storeSectionPrefixes | Select-Object -Unique)
		}

		$BestXchByBaseScaleCode = @{ }
		$sourceRank = 0
		foreach ($candidateIni in @($XchNodesIniCandidates | Select-Object -Unique))
		{
			$sourceRank++
			$XchNodesIniPathsUsed += $candidateIni
			try
			{
				$iniRaw = Get-Content -LiteralPath $candidateIni -Raw -ErrorAction Stop
				foreach ($sectionMatch in [regex]::Matches($iniRaw, '(?ms)^\[(?<name>[^\]]+)\]\s*(?<body>.*?)(?=^\[|\z)'))
				{
					$sectionKey = [string]$sectionMatch.Groups['name'].Value.Trim()
					if ([string]::IsNullOrWhiteSpace($sectionKey)) { continue }

					$currentMap = @{ }
					foreach ($bodyLine in ([string]$sectionMatch.Groups['body'].Value -split "`r?`n"))
					{
						$bodyTrim = $bodyLine.Trim()
						if ([string]::IsNullOrWhiteSpace($bodyTrim) -or $bodyTrim -match '^[;#]') { continue }
						if ($bodyTrim -match '^\s*([^=]+?)\s*=\s*(.*)\s*$') { $currentMap[$matches[1].Trim()] = $matches[2] }
					}
					$get = { param ($k) if ($currentMap.ContainsKey($k)) { [string]$currentMap[$k] } else { $null } }

					$sectionTerminal = $null
					if ($sectionKey -match '(\d{3})$' -and $matches[1] -match '^8\d\d$') { $sectionTerminal = $matches[1] }

					# NUMBER=1 -> 801, NUMBER=801 -> 801
					$numberRaw = & $get 'NUMBER'
					if ($numberRaw) { $numberRaw = $numberRaw.Trim() }
					$scaleCodeFromNumber = $null
					if ($numberRaw -match '^\d+$')
					{
						try
						{
							$n = [int]$numberRaw
							if ($n -ge 0 -and $n -le 99) { $scaleCodeFromNumber = ('8{0:00}' -f $n) }
							elseif ($n -ge 800 -and $n -le 899) { $scaleCodeFromNumber = ('{0:000}' -f $n) }
						}
						catch { }
					}

					# Match rank: 0 = <store><terminal> section, 1 = terminal suffix only, 2 = NUMBER
					$matchedScaleCode = $null
					$matchRank = 99
					if ($sectionTerminal -and $AllowedScaleCodeSet.ContainsKey($sectionTerminal))
					{
						$matchedScaleCode = $sectionTerminal
						$matchRank = 1
						foreach ($prefix in $storeSectionPrefixes) { if ($sectionKey -ieq "$prefix$sectionTerminal") { $matchRank = 0; break } }
					}
					elseif ($scaleCodeFromNumber -and $AllowedScaleCodeSet.ContainsKey($scaleCodeFromNumber))
					{
						$matchedScaleCode = $scaleCodeFromNumber
						$matchRank = 2
					}
					if (-not $matchedScaleCode) { continue }

					# Port=IP 192.168.1.50:3001 -> FullIP / TcpPort / IPNetwork + IPDevice
					$portRaw = & $get 'Port'
					$tcpPort = $null
					$fullIP = $null
					$ipNetwork = $null
					$ipDevice = $null
					if (-not [string]::IsNullOrWhiteSpace($portRaw))
					{
						$portValue = ($portRaw -replace '(?i)^\s*IP\s*', '').Trim()
						if ($portValue -match '^\s*([^:]+?)\s*:(\d+)\s*$')
						{
							$fullIP = $matches[1].Trim()
							$tcpPort = $matches[2].Trim()
						}
						else { $fullIP = $portValue }

						if ($fullIP -match '^\d{1,3}(?:\.\d{1,3}){3}$')
						{
							$lastDot = $fullIP.LastIndexOf('.')
							$ipNetwork = $fullIP.Substring(0, $lastDot + 1)
							$ipDevice = $fullIP.Substring($lastDot + 1)
						}
					}

					$completeness = 0
					if ($fullIP) { $completeness += 10 }
					if (-not [string]::IsNullOrWhiteSpace((& $get 'Scalevendor'))) { $completeness += 5 }
					if (-not [string]::IsNullOrWhiteSpace((& $get 'ScaleModel'))) { $completeness += 5 }
					if ($portRaw) { $completeness += 3 }
					foreach ($k in 'Baud', 'Type', 'BaseMode') { if (-not [string]::IsNullOrWhiteSpace((& $get $k))) { $completeness += 1 } }

					$entryObj = [PSCustomObject]@{
						SourceRank   = $sourceRank
						MatchRank    = $matchRank
						Completeness = $completeness
						Section      = $sectionKey
						BaseMode     = & $get 'BaseMode'
						Type         = & $get 'Type'
						ScaleBrand   = & $get 'Scalevendor'
						ScaleModel   = & $get 'ScaleModel'
						Baud         = & $get 'Baud'
						PortRaw      = $portRaw
						TcpPort      = $tcpPort
						IPNetwork    = $ipNetwork
						IPDevice     = $ipDevice
						FullIP       = $fullIP
						Sub          = & $get 'Sub'
						PlumDir      = & $get 'PlumDir'
					}

					# Keep the best section: lower MatchRank, then more complete, then earlier file
					$existingXch = $BestXchByBaseScaleCode[$matchedScaleCode]
					if (-not $existingXch -or
						$entryObj.MatchRank -lt $existingXch.MatchRank -or
						($entryObj.MatchRank -eq $existingXch.MatchRank -and $entryObj.Completeness -gt $existingXch.Completeness) -or
						($entryObj.MatchRank -eq $existingXch.MatchRank -and $entryObj.Completeness -eq $existingXch.Completeness -and $entryObj.SourceRank -lt $existingXch.SourceRank))
					{
						$BestXchByBaseScaleCode[$matchedScaleCode] = $entryObj
					}
				}
			}
			catch
			{
				Write_Log "Failed to read XchNodes.ini '$candidateIni': $_" "yellow"
			}
		}

		# Fill the TER scales' empty fields from their best XchNodes section
		foreach ($terCode in @($TerScaleBy80X.Keys))
		{
			$xchEntry = $BestXchByBaseScaleCode[$terCode]
			$terEntry = $TerScaleBy80X[$terCode]
			if (-not $xchEntry -or -not $terEntry) { continue }
			foreach ($field in 'IPNetwork', 'IPDevice', 'FullIP', 'PortRaw', 'TcpPort', 'Section', 'BaseMode', 'Type', 'Baud', 'Sub', 'PlumDir', 'ScaleModel')
			{
				if ([string]::IsNullOrWhiteSpace([string]$terEntry.$field) -and $xchEntry.$field) { $terEntry.$field = [string]$xchEntry.$field }
			}
			if ([string]::IsNullOrWhiteSpace([string]$terEntry.ScaleBrand) -and $xchEntry.ScaleBrand)
			{
				$terEntry.ScaleBrand = [string]$xchEntry.ScaleBrand
				$terEntry.Scalevendor = [string]$xchEntry.ScaleBrand
			}
		}
	}

	# ====================================================================================
	# 2.75 DROP TER SCALES THAT ARE ALSO A TBS SCALE (same IP) - the TBS row wins
	# ====================================================================================
	$normalizedIp = {
		param ($e)
		if (-not [string]::IsNullOrWhiteSpace([string]$e.FullIP)) { return ([string]$e.FullIP).Trim().ToLowerInvariant() }
		if (-not [string]::IsNullOrWhiteSpace([string]$e.IPNetwork) -or -not [string]::IsNullOrWhiteSpace([string]$e.IPDevice))
		{
			return ("$([string]$e.IPNetwork)$([string]$e.IPDevice)").Trim().ToLowerInvariant()
		}
		return $null
	}
	$TbsByNormalizedIp = @{ }
	foreach ($tbsEntry in @($TbsScaleRows))
	{
		$ip = & $normalizedIp $tbsEntry
		if ($ip -and -not $TbsByNormalizedIp.ContainsKey($ip)) { $TbsByNormalizedIp[$ip] = $tbsEntry }
	}

	$FinalTerBy80X = @{ }
	foreach ($terCode in @($TerScaleBy80X.Keys | Sort-Object { [int]$_ }))
	{
		$terEntry = $TerScaleBy80X[$terCode]
		$ip = & $normalizedIp $terEntry
		if ($ip -and $TbsByNormalizedIp.ContainsKey($ip))
		{
			$matchingTbsRow = $TbsByNormalizedIp[$ip]
			foreach ($field in 'Path', 'HostPath', 'PortRaw', 'TcpPort', 'Section', 'BaseMode', 'Type', 'Baud', 'Sub', 'PlumDir')
			{
				if ([string]::IsNullOrWhiteSpace([string]$matchingTbsRow.$field) -and $terEntry.$field) { $matchingTbsRow.$field = [string]$terEntry.$field }
			}
			continue
		}
		$FinalTerBy80X[$terCode] = $terEntry
	}

	# ====================================================================================
	# 2.9 FINAL SCALE KEYS AND MAPS (TER by 80X first, then TBS by original ScaleCode)
	# ====================================================================================
	$addScale = {
		param ([string]$finalKey, $entry, [string]$source)

		$finalObj = [PSCustomObject]@{
			Code          = $finalKey
			ScaleCode     = $finalKey
			BaseScaleCode = $entry.BaseScaleCode
			Label         = $entry.Label
			ScaleName     = $entry.ScaleName
			Path          = $entry.Path
			HostPath      = $entry.HostPath
			ScaleBrand    = $entry.ScaleBrand
			Scalevendor   = if ($entry.Scalevendor) { $entry.Scalevendor } else { $entry.ScaleBrand }
			ScaleModel    = $entry.ScaleModel
			ScaleLocation = $entry.ScaleLocation
			IPNetwork     = $entry.IPNetwork
			IPDevice      = $entry.IPDevice
			FullIP        = $entry.FullIP
			PortRaw       = $entry.PortRaw
			TcpPort       = $entry.TcpPort
			Section       = $entry.Section
			BaseMode      = $entry.BaseMode
			Type          = $entry.Type
			Baud          = $entry.Baud
			Sub           = $entry.Sub
			PlumDir       = $entry.PlumDir
			Active        = $entry.Active
			Source        = $source
		}

		$ScaleCodes += $finalKey
		$ScaleCodeToIPInfo[$finalKey] = $finalObj
		$ScaleLabels[$finalKey] = $finalObj.Label
		$ScaleExePaths[$finalKey] = $finalObj.Path
		$ScaleCodeToPath[$finalKey] = $finalObj.Path
		if ($finalObj.ScaleName)
		{
			$ScaleNameToCode[$finalObj.ScaleName] = $finalKey
			$ScaleNameToPath[$finalObj.ScaleName] = $finalObj.Path
			$ScaleCodeToName[$finalKey] = $finalObj.ScaleName
		}
		foreach ($fm in $scaleFieldMaps)
		{
			$fv = $finalObj.($fm[0])
			if ($fv) { (Get-Variable -Name $fm[1] -ValueOnly)[$finalKey] = $fv }
		}
	}

	foreach ($terCode in @($FinalTerBy80X.Keys | Sort-Object { [int]$_ }))
	{
		. $addScale ([string]$terCode) $FinalTerBy80X[$terCode] 'TER'
	}

	# A TBS ScaleCode that collides with an existing key gets a _TBS<n> suffix
	$TbsKeyCounts = @{ }
	foreach ($tbsEntry in @($TbsScaleRows))
	{
		$rawKey = [string]$tbsEntry.RawScaleCode
		if ([string]::IsNullOrWhiteSpace($rawKey)) { continue }
		$finalKey = $rawKey
		if (-not $TbsKeyCounts.ContainsKey($rawKey)) { $TbsKeyCounts[$rawKey] = 1 }
		while ($ScaleCodeToIPInfo.ContainsKey($finalKey))
		{
			$TbsKeyCounts[$rawKey]++
			$finalKey = "{0}_TBS{1}" -f $rawKey, $TbsKeyCounts[$rawKey]
		}
		. $addScale $finalKey $tbsEntry 'TBS'
	}

	# ====================================================================================
	# 3. REVERSE MAPS, COUNTS, BIZERBA (WINDOWS) SCALES
	# ====================================================================================
	foreach ($kv in @($LaneNumToMachineName.GetEnumerator())) { $MachineNameToLaneNum[$kv.Value] = $kv.Key }
	foreach ($kv in @($BackofficeNumToMachineName.GetEnumerator()))
	{
		$MachineNameToBackofficeNum[$kv.Value] = $kv.Key
		if ($BackofficeNumToPath.ContainsKey($kv.Key)) { $MachineNameToBackofficePath[$kv.Value] = $BackofficeNumToPath[$kv.Key] }
	}

	$NumberOfLanes = $LaneMachineNames.Count
	$NumberOfServers = if ($ServerMachineName) { 1 } else { 0 }
	$NumberOfBackoffices = $BackofficeNumToMachineName.Count
	$NumberOfScales = $ScaleCodes.Count

	# Bizerba scales run Windows (C$ shares, credentials, file deployment)
	foreach ($code in @($ScaleCodeToIPInfo.Keys))
	{
		$scale = $ScaleCodeToIPInfo[$code]
		foreach ($field in 'ScaleBrand', 'Scalevendor', 'ScaleModel', 'ScaleName', 'Label')
		{
			if ($scale.$field -and ([string]$scale.$field -match '(?i)bizerba')) { $WindowsScales[$code] = $scale; break }
		}
	}

	. $publish

	# ====================================================================================
	# 4. CACHE WRITE (keys sorted so the file is readable and stable)
	# ====================================================================================
	try
	{
		$numKey = { [double]((([string]$_.Key) -replace '[^\d]', '0')) }
		$sortMap = {
			param ($Map, [object[]]$SortBy)
			$o = [ordered]@{ }
			foreach ($kv in @($Map.GetEnumerator() | Sort-Object $SortBy)) { $o[$kv.Key] = $kv.Value }
			return $o
		}

		$out = [ordered]@{ }
		foreach ($n in $mapNames) { $out[$n] = & $sortMap (Get-Variable -Name $n -ValueOnly) @($numKey, { $_.Key }) }

		# Lanes: only 3-digit lane numbers; the per-machine maps follow the lane order
		$sortedLaneNums = @($LaneNumToMachineName.Keys | Where-Object { $_ -match '^\d{3}$' } | Sort-Object { [int]$_ })
		$out['LaneNumToMachineName'] = [ordered]@{ }
		foreach ($k in 'LaneMachineLabels', 'LaneMachinePath', 'LaneMachineToServerPath') { $out[$k] = [ordered]@{ } }
		foreach ($ln in $sortedLaneNums)
		{
			$m = $LaneNumToMachineName[$ln]
			$out['LaneNumToMachineName'][$ln] = $m
			if (-not $m) { continue }
			foreach ($k in 'LaneMachineLabels', 'LaneMachinePath', 'LaneMachineToServerPath')
			{
				$src = Get-Variable -Name $k -ValueOnly
				if ($src.ContainsKey($m)) { $out[$k][$m] = $src[$m] }
			}
		}
		$byValueThenKey = @({ [double]((([string]$_.Value) -replace '[^\d]', '0')) }, { $_.Key })
		$out['MachineNameToLaneNum'] = & $sortMap $MachineNameToLaneNum $byValueThenKey
		$out['MachineNameToBackofficeNum'] = & $sortMap $MachineNameToBackofficeNum $byValueThenKey
		$out['ScaleNameToCode'] = & $sortMap $ScaleNameToCode @({ $_.Value }, { $_.Key })
		$out['ScaleNameToPath'] = & $sortMap $ScaleNameToPath @({ $_.Key })
		$out['MachineNameToBackofficePath'] = & $sortMap $MachineNameToBackofficePath @({ $_.Key })

		$outLaneMachineNames = @($LaneMachineNames)
		if ($outLaneMachineNames.Count -gt 0 -and $MachineNameToLaneNum.Count -gt 0)
		{
			$outLaneMachineNames = @($outLaneMachineNames | Sort-Object { [double](([string]$MachineNameToLaneNum[$_]) -replace '[^\d]', '0') }, { $_ })
		}

		$results = [ordered]@{
			NumberOfLanes       = $NumberOfLanes
			NumberOfServers     = $NumberOfServers
			NumberOfBackoffices = $NumberOfBackoffices
			NumberOfScales      = $NumberOfScales
			LaneMachineNames    = $outLaneMachineNames
			ScaleCodes          = @($ScaleCodes | Sort-Object { [double](([string]$_) -replace '[^\d]', '0') }, { $_ })
			ServerMachineName   = $ServerMachineName
			ServerLabel         = $ServerLabel
			ServerPath          = $ServerPath
		}
		foreach ($n in $mapNames) { $results[$n] = $out[$n] }
		$results['XchNodesIniPathsUsed'] = @($XchNodesIniPathsUsed | Select-Object -Unique)

		$cacheOut = [ordered]@{
			LastUpdated        = (Get-Date).ToString('o')
			CacheSchemaVersion = $CacheSchemaVersion
			StoreNumber        = $StoreNumberNorm
			Results            = $results
		}
		($cacheOut | ConvertTo-Json -Depth 16) | Set-Content -LiteralPath $CachePath -Encoding UTF8 -Force
	}
	catch { }

	return $Nodes
}

# ===================================================================================================
#                              FUNCTION: Show_Node_Selection_Form
# ---------------------------------------------------------------------------------------------------
# Description:
#   The shared node picker: tabbed WinForms dialog for Lanes, Scales, Backoffices, or any
#   combination. Returns a hashtable with keys matching the selected node types (Lanes,
#   Scales, Backoffices), or $null when the user cancels or selects nothing.
#
# Parameters:
#   - StoreNumber        [string] (required): store to look up
#   - NodeTypes          [string[]]: "Lane" / "Scale" / "Backoffice" (default: all three)
#   - OnlyBizerbaScales  [switch]: limit the Scales tab to Bizerba scales
#   - Title              [string]: dialog title (default "Select Nodes to Process")
#   - ExcludedNodes      [string[]]: machine names and/or 3-digit lane numbers to hide
#   - SingleLaneOnly     [switch]: cap lane selection at 1
#   - LaneSelectionLimit [int]: cap lane selection at N (0 = unlimited)
# ===================================================================================================

function Show_Node_Selection_Form
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter()]
		[ValidateSet("Lane", "Scale", "Backoffice")]
		[string[]]$NodeTypes = @("Lane", "Scale", "Backoffice"),
		[switch]$OnlyBizerbaScales,
		[string]$Title = "Select Nodes to Process",
		[string[]]$ExcludedNodes,
		[switch]$SingleLaneOnly,
		# NEW: cap lane selection at 1 (optional)
		[int]$LaneSelectionLimit = 0 # NEW: cap lane selection at N (0 = unlimited)
	)

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	# The "at most N lanes" notice shows once per dialog (a cancelled dialog used to keep it muted)
	Remove-Variable -Name __laneLimitWarned -Scope Script -ErrorAction SilentlyContinue

	# Normalize exclusions (names lowercased; lanes as 3-digit)
	$__ex_names = @()
	$__ex_lanes = @()
	if ($ExcludedNodes)
	{
		foreach ($e in $ExcludedNodes)
		{
			if ($null -ne $e -and "$e".Trim())
			{
				$s = "$e".Trim()
				if ($s -match '(\d{1,3})$')
				{
					try { $__ex_lanes += ('{0:D3}' -f ([int]$matches[1])) }
					catch { }
				}
				$__ex_names += $s.ToLower()
			}
		}
		$__ex_names = $__ex_names | Select-Object -Unique
		$__ex_lanes = $__ex_lanes | Select-Object -Unique
	}

	$form = New-Object System.Windows.Forms.Form
	$form.Text = $Title
	$form.Size = New-Object System.Drawing.Size(430, 450)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	$tabs = New-Object System.Windows.Forms.TabControl
	$tabs.Location = New-Object System.Drawing.Point(10, 10)
	$tabs.Size = New-Object System.Drawing.Size(390, 320)
	$form.Controls.Add($tabs)

	$tabControls = @{ }

	# ----- Lanes Tab -----
	if ("Lane" -in $NodeTypes)
	{
		$tabLanes = New-Object System.Windows.Forms.TabPage
		$tabLanes.Text = "Lanes"
		$clbLanes = New-Object System.Windows.Forms.CheckedListBox
		$clbLanes.Location = New-Object System.Drawing.Point(10, 10)
		$clbLanes.Size = New-Object System.Drawing.Size(350, 270)
		$clbLanes.CheckOnClick = $true
		$tabLanes.Controls.Add($clbLanes)
		$tabs.TabPages.Add($tabLanes)
		$tabControls["Lanes"] = $clbLanes

		# Load lane node data from FunctionResults
		$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
		$LaneMachineLabels = $script:FunctionResults['LaneMachineLabels']
		$LaneMachinePath = $script:FunctionResults['LaneMachinePath']
		$LaneMachineToServerPath = $script:FunctionResults['LaneMachineToServerPath']

		$allLanes = @()
		if ($script:FunctionResults.ContainsKey('LaneMachineNames') -and $script:FunctionResults['LaneMachineNames'].Count -gt 0)
		{
			$allLanes = $script:FunctionResults['LaneMachineNames']
		}
		elseif (Test-Path -Path $OfficePath)
		{
			$laneFolders = Get-ChildItem -Path $OfficePath -Directory -Filter "XF${StoreNumber}0*"
			if ($laneFolders) { $allLanes = $laneFolders | ForEach-Object { $_.Name.Substring($_.Name.Length - 3, 3) } }
		}

		# machine -> 3-digit lane number (first match, as before)
		$machineToLaneNum = @{ }
		if ($LaneNumToMachineName)
		{
			foreach ($key in @($LaneNumToMachineName.Keys | Sort-Object))
			{
				$m = [string]$LaneNumToMachineName[$key]
				if ($key -match '^\d{3}$' -and $m -and -not $machineToLaneNum.ContainsKey($m)) { $machineToLaneNum[$m] = $key }
			}
		}

		# Lane order (POS2 before POS10), machines without a lane number last
		foreach ($laneMachine in ($allLanes | Sort-Object { $n = $machineToLaneNum[[string]$_]; if ($n) { [int]$n } else { 99999 } }, { [string]$_ }))
		{
			$laneNum = $machineToLaneNum[[string]$laneMachine]
			# XF-folder fallback lists 3-digit lane numbers, not machine names
			if (-not $laneNum -and [string]$laneMachine -match '^\d{3}$') { $laneNum = [string]$laneMachine }

			# Skip excluded by machine name or lane number
			$__skip = $false
			if ($laneMachine -and ($__ex_names -contains $laneMachine.ToLower())) { $__skip = $true }
			elseif ($laneNum -and ($__ex_lanes -contains $laneNum)) { $__skip = $true }
			if ($__skip) { continue }

			# Add item
			$obj = [PSCustomObject]@{
				LaneNumber  = $laneNum
				MachineName = $laneMachine
				Label	    = $LaneMachineLabels[$laneMachine]
				Path	    = $LaneMachinePath[$laneMachine]
				ServerPath  = $LaneMachineToServerPath[$laneMachine]
				DisplayName = if ($laneNum) { "$laneMachine ($laneNum)" } else { $laneMachine }
			}
			$obj | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.DisplayName } -Force
			$clbLanes.Items.Add($obj) | Out-Null
		}

		# --- Enforce optional selection cap for lanes (single or N) ---
		$__laneCap = if ($SingleLaneOnly) { 1 }
		elseif ($LaneSelectionLimit -gt 0) { [int]$LaneSelectionLimit }
		else { 0 }
		if ($__laneCap -gt 0)
		{
			$clbLanes.Add_ItemCheck({
					# If trying to check and it would exceed the cap, cancel the check.
					if ($_.NewValue -ne [System.Windows.Forms.CheckState]::Checked) { return }
					$idx = $_.Index
					$checkedCount = 0
					for ($i = 0; $i -lt $clbLanes.Items.Count; $i++)
					{
						if ($i -ne $idx -and $clbLanes.GetItemChecked($i)) { $checkedCount++ }
					}
					if (($checkedCount + 1) -gt $__laneCap)
					{
						$_.NewValue = [System.Windows.Forms.CheckState]::Unchecked
						if (-not $script:__laneLimitWarned)
						{
							[System.Windows.Forms.MessageBox]::Show("You can select at most $__laneCap lane(s).", "Selection limit",
								[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
							$script:__laneLimitWarned = $true
						}
					}
				})
		}
	}

	# ----- Scales Tab -----
	if ("Scale" -in $NodeTypes)
	{
		$tabScales = New-Object System.Windows.Forms.TabPage
		$tabScales.Text = "Scales"
		$clbScales = New-Object System.Windows.Forms.CheckedListBox
		$clbScales.Location = New-Object System.Drawing.Point(10, 10)
		$clbScales.Size = New-Object System.Drawing.Size(350, 270)
		$clbScales.CheckOnClick = $true
		$tabScales.Controls.Add($clbScales)
		$tabs.TabPages.Add($tabScales)
		$tabControls["Scales"] = $clbScales

		$allScales = @()
		if ($script:FunctionResults.ContainsKey('ScaleCodeToIPInfo'))
		{
			$allScales = $script:FunctionResults['ScaleCodeToIPInfo'].Values
			if ($OnlyBizerbaScales) { $allScales = $allScales | Where-Object { $_.ScaleBrand -match 'bizerba' } }
		}
		# numeric order; non-numeric TBS codes (e.g. 801_TBS2) sort by their digits
		$sortedScales = $allScales | Sort-Object { [double](([string]$_.ScaleCode -replace '[^\d]', '0')) }, { [string]$_.ScaleCode }

		$uniqueScaleIPs = @{ }
		$dedupedScales = @()
		foreach ($scale in $sortedScales)
		{
			$ipKey = if ($scale.FullIP) { $scale.FullIP }
			elseif ($scale.IPNetwork -and $scale.IPDevice) { (([string]$scale.IPNetwork).TrimEnd('.') + '.' + [string]$scale.IPDevice) }
			else { $null }
			if ($ipKey -and -not $uniqueScaleIPs.ContainsKey($ipKey))
			{
				$uniqueScaleIPs[$ipKey] = $true
				$dedupedScales += $scale
			}
		}

		foreach ($scale in $dedupedScales)
		{
			# Skip excluded by name/code
			$__skip = $false
			if ($scale.ScaleName -and ($__ex_names -contains ($scale.ScaleName.ToLower()))) { $__skip = $true }
			elseif ([string]$scale.ScaleCode -match '^\d+$' -and ($__ex_lanes -contains ('{0:D3}' -f ([int]$scale.ScaleCode)))) { $__skip = $true }
			if ($__skip) { continue }

			$ip = if ($scale.IPNetwork -and $scale.IPDevice) { (([string]$scale.IPNetwork).TrimEnd('.') + '.' + [string]$scale.IPDevice) }
			elseif ($scale.FullIP) { [string]$scale.FullIP }   # TER scales known by host name only
			else { "" }
			$displayName = "$($scale.ScaleName) [$ip]"
			$scaleObj = [PSCustomObject]@{
				ScaleCode   = $scale.ScaleCode
				ScaleName   = $scale.ScaleName
				IPAddress   = $ip
				Vendor	    = $scale.Scalevendor
				FullIP	    = $scale.FullIP
				IPNetwork   = $scale.IPNetwork
				IPDevice    = $scale.IPDevice
				ScaleBrand  = $scale.ScaleBrand
				DisplayName = $displayName
			}
			$scaleObj | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.DisplayName } -Force
			$clbScales.Items.Add($scaleObj) | Out-Null
		}
	}

	# ----- Backoffices Tab -----
	if ("Backoffice" -in $NodeTypes)
	{
		$tabBO = New-Object System.Windows.Forms.TabPage
		$tabBO.Text = "Backoffices"
		$clbBO = New-Object System.Windows.Forms.CheckedListBox
		$clbBO.Location = New-Object System.Drawing.Point(10, 10)
		$clbBO.Size = New-Object System.Drawing.Size(350, 270)
		$clbBO.CheckOnClick = $true
		$tabBO.Controls.Add($clbBO)
		$tabs.TabPages.Add($tabBO)
		$tabControls["Backoffices"] = $clbBO

		$boDict = if ($script:FunctionResults.ContainsKey('BackofficeNumToMachineName')) { $script:FunctionResults['BackofficeNumToMachineName'] }
		else { @{ } }
		$boLabels = $script:FunctionResults['BackofficeNumToLabel']
		$boPaths = $script:FunctionResults['BackofficeNumToPath']

		$seenBonumbers = @{ }
		foreach ($boNumKey in $boDict.Keys | Sort-Object)
		{
			if ($boNumKey -match '(\d{3})')
			{
				$bonum = $matches[1]

				# Skip excluded by machine name or BO number
				$__skip = $false
				$machineName = $boDict[$boNumKey]
				if ($machineName -and ($__ex_names -contains $machineName.ToLower())) { $__skip = $true }
				elseif ($__ex_lanes -contains $bonum) { $__skip = $true }
				if ($__skip) { continue }

				if (-not $seenBonumbers.ContainsKey($bonum))
				{
					$seenBonumbers[$bonum] = $true
					$label = $boLabels[$boNumKey]
					$path = $boPaths[$boNumKey]

					$obj = [PSCustomObject]@{
						BONumber    = $bonum
						MachineName = $machineName
						Label	    = $label
						Path	    = $path
						DisplayName = if ($machineName) { "$machineName ($bonum)" } else { "Unknown ($bonum)" }
					}
					$obj | Add-Member ScriptMethod ToString { $this.DisplayName } -Force
					$clbBO.Items.Add($obj) | Out-Null
				}
			}
		}
	}

	$btnSelectAll = New-Object System.Windows.Forms.Button
	$btnSelectAll.Location = New-Object System.Drawing.Point(20, 340)
	$btnSelectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnSelectAll.Text = "Select All"
	$btnSelectAll.BackColor = [System.Drawing.SystemColors]::Control

	$btnDeselectAll = New-Object System.Windows.Forms.Button
	$btnDeselectAll.Location = New-Object System.Drawing.Point(220, 340)
	$btnDeselectAll.Size = New-Object System.Drawing.Size(180, 32)
	$btnDeselectAll.Text = "Deselect All"
	$btnDeselectAll.BackColor = [System.Drawing.SystemColors]::Control

	$setBtnColor = {
		param ($btn,
			$state)
		switch ($state)
		{
			1 { $btn.BackColor = [System.Drawing.Color]::Yellow }
			2 { $btn.BackColor = [System.Drawing.Color]::LightGreen }
			Default { $btn.BackColor = [System.Drawing.SystemColors]::Control }
		}
	}
	&$setBtnColor $btnSelectAll 0
	&$setBtnColor $btnDeselectAll 0

	# --- SINGLE TAB SELECT ALL LOGIC ---
	if ($tabs.TabPages.Count -eq 1)
	{
		$clb = $tabControls[$tabs.TabPages[0].Text]
		$isAllChecked = {
			for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { return $false } }
			return $clb.Items.Count -gt 0
		}
		$isAnyChecked = {
			for ($i = 0; $i -lt $clb.Items.Count; $i++) { if ($clb.GetItemChecked($i)) { return $true } }
			return $false
		}
		$btnSelectAll.Add_Click({
				$allChecked = & $isAllChecked
				if (-not $allChecked)
				{
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }
					& $setBtnColor $btnSelectAll 2
				}
			})
		$btnDeselectAll.Add_Click({
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }
				& $setBtnColor $btnSelectAll 0
				& $setBtnColor $btnDeselectAll 0
			})
		# (MouseUp / KeyUp only: ItemCheck fires before the state changes, and a sleep in it made
		#  Select All stall ~30 ms per item)
		foreach ($event in @("Add_MouseUp", "Add_KeyUp"))
		{
			$clb.$event.Invoke({
					$allChecked = & $isAllChecked
					$anyChecked = & $isAnyChecked
					if ($allChecked) { & $setBtnColor $btnSelectAll 2 }
					elseif ($anyChecked) { & $setBtnColor $btnSelectAll 1 }
					else { & $setBtnColor $btnSelectAll 0 }
				})
		}
	}
	else
	{
		# ---- NORMAL MULTI-TAB LOGIC ----
		$tabSelectState = @{ }
		$lastSelectTabIndex = $null
		$selectAllYellowTabIndex = $null

		$btnSelectAll.Add_Click({
				$tabName = $tabs.SelectedTab.Text
				$clb = $tabControls[$tabName]
				$tabIndex = $tabs.SelectedIndex
				$currentTabAllChecked = $true
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { $currentTabAllChecked = $false; break } }
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked)
				{
					foreach ($k in $tabControls.Keys)
					{
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
						if ($tabIndex2 -eq -1) { continue }
						$tabSelectState[$tabIndex2] = 0
					}
					$tabSelectState[$tabIndex] = 0
					$selectAllYellowTabIndex = $null
					&$setBtnColor $btnSelectAll 0
					return
				}
				if ($currentTabAllChecked -and -not $allTabsChecked)
				{
					foreach ($k in $tabControls.Keys)
					{
						$list = $tabControls[$k]
						for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $true) }
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
						if ($tabIndex2 -eq -1) { continue }
						$tabSelectState[$tabIndex2] = 2
					}
					&$setBtnColor $btnSelectAll 2
					$selectAllYellowTabIndex = $null
					return
				}
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) }
				$tabSelectState[$tabIndex] = 1
				$lastSelectTabIndex = $tabIndex
				$selectAllYellowTabIndex = $tabIndex
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
				else { & $setBtnColor $btnSelectAll 1 }
			})
		$btnDeselectAll.Add_Click({
				$tabName = $tabs.SelectedTab.Text
				$clb = $tabControls[$tabName]
				$tabIndex = $tabs.SelectedIndex
				$noneChecked = $true
				for ($i = 0; $i -lt $clb.Items.Count; $i++) { if ($clb.GetItemChecked($i)) { $noneChecked = $false; break } }
				if ($noneChecked)
				{
					$originalTab = $tabs.SelectedTab
					foreach ($k in $tabControls.Keys)
					{
						$tabIndex2 = -1
						for ($t = 0; $t -lt $tabs.TabPages.Count; $t++) { if ($tabs.TabPages[$t].Text -eq $k) { $tabIndex2 = $t; break } }
						if ($tabIndex2 -eq -1) { continue }
						$tabs.SelectedTab = $tabs.TabPages[$tabIndex2]
						$list = $tabControls[$k]
						for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) }
						$tabSelectState[$tabIndex2] = 0
						$list.Refresh()
					}
					$tabs.SelectedTab = $originalTab
				}
				else
				{
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) }
					$tabSelectState[$tabIndex] = 0
				}
				&$setBtnColor $btnDeselectAll 0
				&$setBtnColor $btnSelectAll 0
				$selectAllYellowTabIndex = $null
			})
		foreach ($clb in $tabControls.Values)
		{
			$clb.Add_MouseUp({
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
					}
					if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++) { if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break } }
						if ($allChecked -and $clbLocal.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
						else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
					}
				})
			$clb.Add_KeyUp({
					$allTabsChecked = $true
					foreach ($clbTest in $tabControls.Values)
					{
						for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
					}
					if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
					else
					{
						$tabIndex = $tabs.SelectedIndex
						$clbLocal = $tabControls[$tabs.SelectedTab.Text]
						$allChecked = $true
						for ($i = 0; $i -lt $clbLocal.Items.Count; $i++) { if (-not $clbLocal.GetItemChecked($i)) { $allChecked = $false; break } }
						if ($allChecked -and $clbLocal.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
						else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
					}
				})
		}
		$tabs.add_SelectedIndexChanged({
				$tabIndex = $tabs.SelectedIndex
				$clb = $tabControls[$tabs.SelectedTab.Text]
				$allTabsChecked = $true
				foreach ($clbTest in $tabControls.Values)
				{
					for ($i = 0; $i -lt $clbTest.Items.Count; $i++) { if (-not $clbTest.GetItemChecked($i)) { $allTabsChecked = $false; break } }
				}
				if ($allTabsChecked) { & $setBtnColor $btnSelectAll 2; $selectAllYellowTabIndex = $null }
				else
				{
					$allChecked = $true
					for ($i = 0; $i -lt $clb.Items.Count; $i++) { if (-not $clb.GetItemChecked($i)) { $allChecked = $false; break } }
					if ($allChecked -and $clb.Items.Count -gt 0) { & $setBtnColor $btnSelectAll 1; $selectAllYellowTabIndex = $tabIndex }
					else { & $setBtnColor $btnSelectAll 0; $selectAllYellowTabIndex = $null }
				}
				&$setBtnColor $btnDeselectAll 0
			})
	}

	$form.Controls.Add($btnSelectAll)
	$form.Controls.Add($btnDeselectAll)

	# OK Button
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Location = New-Object System.Drawing.Point(60, 380)
	$btnOK.Size = New-Object System.Drawing.Size(140, 32)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $btnOK
	$form.Controls.Add($btnOK)

	# Cancel Button
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(220, 380)
	$btnCancel.Size = New-Object System.Drawing.Size(140, 32)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $btnCancel
	$form.Controls.Add($btnCancel)
	if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $form }

	# ----- Show form & collect selections (the form is disposed afterwards: it leaked a window
	#       handle set per use) -----
	$dialogResult = $form.ShowDialog()
	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		$form.Dispose()
		return $null
	}

	$result = @{ }
	if ("Lane" -in $NodeTypes)
	{
		$clb = $tabControls["Lanes"]
		$selectedLaneNums = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['LaneNumber'] -and $item.LaneNumber)
				{
					$selectedLaneNums += $item.LaneNumber
				}
			}
		}
		$result.Lanes = $selectedLaneNums
	}
	if ("Scale" -in $NodeTypes)
	{
		$clb = $tabControls["Scales"]
		$selectedScaleCodes = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['ScaleCode'] -and $item.ScaleCode)
				{
					$selectedScaleCodes += "$($item.ScaleCode)"
				}
			}
		}
		$result.Scales = $selectedScaleCodes
	}
	if ("Backoffice" -in $NodeTypes)
	{
		$clb = $tabControls["Backoffices"]
		$selectedBONums = @()
		for ($i = 0; $i -lt $clb.Items.Count; $i++)
		{
			if ($clb.GetItemChecked($i))
			{
				$item = $clb.Items[$i]
				if ($item.PSObject.Properties['BONumber'] -and $item.BONumber)
				{
					$selectedBONums += $item.BONumber
				}
			}
		}
		$result.Backoffices = $selectedBONums
	}

	$form.Dispose()

	# Basic validation
	if ($result.Keys.Count -eq 0 -or ($result.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum -eq 0)
	{
		# [void]: the MessageBox result ('OK') must not become this function's return value
		[void][System.Windows.Forms.MessageBox]::Show("No nodes selected.", "Information",
			[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
		Remove-Variable -Name __laneLimitWarned -Scope Script -ErrorAction SilentlyContinue
		return $null
	}

	Remove-Variable -Name __laneLimitWarned -Scope Script -ErrorAction SilentlyContinue
	return $result
}

# ===================================================================================================
#                              FUNCTION: Delete_DBS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Deletes specific file types (.txt and .dwr) from the Dbs folders of selected lanes in a
#   store, with a prompt to also include .sus files. Lane machine paths come from the
#   Retrieve_Nodes results; the actual deletion is done by Delete_Files per lane.
#
# Parameters:
#   - StoreNumber  [string]: store whose lanes are processed (required)
#
# Variables:
#   - $script:FunctionResults: LaneMachineNames and LaneNumToMachineName (input)
#
# Notes:
#   - Retrieve_Nodes must have been run first.
# ===================================================================================================

function Delete_DBS
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)

	Write_Log "`r`n==================== Starting Delete_DBS Function ====================`r`n" "blue"

	# Check if FunctionResults has the necessary data
	if (-not $script:FunctionResults.ContainsKey('LaneMachineNames') -or
		-not $script:FunctionResults.ContainsKey('LaneNumToMachineName'))
	{
		Write_Log "Lane information is not available. Please run Retrieve_Nodes first." "Red"
		return
	}

	# Retrieve lane information
	$LaneMachineNames = $script:FunctionResults['LaneMachineNames']
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']

	if ($LaneMachineNames.Count -eq 0)
	{
		Write_Log "No lanes found for Store Number: $StoreNumber." "Yellow"
		return
	}

	# Prompt user to include .sus files in deletion
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Delete_DBS Confirmation"
	$form.Size = New-Object System.Drawing.Size(400, 200)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false

	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Do you want to include .sus files in the deletion?"
	$label.AutoSize = $true
	$label.Location = New-Object System.Drawing.Point(20, 30)
	$form.Controls.Add($label)

	$checkboxSus = New-Object System.Windows.Forms.CheckBox
	$checkboxSus.Text = "Include .sus files"
	$checkboxSus.AutoSize = $true
	$checkboxSus.Location = New-Object System.Drawing.Point(20, 60)
	$form.Controls.Add($checkboxSus)

	$buttonOK = New-Object System.Windows.Forms.Button
	$buttonOK.Text = "OK"
	$buttonOK.Location = New-Object System.Drawing.Point(100, 120)
	$buttonOK.Size = New-Object System.Drawing.Size(80, 30)
	$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.Controls.Add($buttonOK)

	$buttonCancel = New-Object System.Windows.Forms.Button
	$buttonCancel.Text = "Cancel"
	$buttonCancel.Location = New-Object System.Drawing.Point(200, 120)
	$buttonCancel.Size = New-Object System.Drawing.Size(80, 30)
	$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($buttonCancel)

	$form.AcceptButton = $buttonOK
	$form.CancelButton = $buttonCancel

	$dialogResult = $form.ShowDialog()

	if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
	{
		Write_Log "User canceled the deletion process." "Yellow"
		return
	}

	$includeSus = $checkboxSus.Checked

	# Define file types to delete
	$fileExtensions = @("*.txt", "*.dwr")
	if ($includeSus)
	{
		$fileExtensions += "*.sus"
	}

	Write_Log "Starting deletion of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "Green"

	# Show the selection dialog to choose lanes
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"

	if (-not $selection)
	{
		# User canceled the dialog
		Write_Log "User canceled the lane selection." "Yellow"
		return
	}

	# Lane numbers (the picker returns 3-digit strings; lane objects are accepted too), mapped lanes only
	$selectedLanes = @($selection.Lanes | ForEach-Object {
			if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber } else { [string]$_ }
		} | Where-Object { $LaneNumToMachineName.ContainsKey($_) })

	if (-not $selectedLanes -or $selectedLanes.Count -eq 0)
	{
		Write_Log "No valid lanes selected for processing." "Yellow"
		return
	}

	# Initialize counters
	$totalDeleted = 0
	$totalFailed = 0

	foreach ($lane in $selectedLanes)
	{
		if ($LaneNumToMachineName.ContainsKey($lane))
		{
			$machineName = $LaneNumToMachineName[$lane]

			if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
			{
				Write_Log "Lane #${lane}: Machine name is invalid or unknown. Skipping deletion." "Yellow"
				continue
			}

			# Construct the target path (modify this path as per your environment)
			$targetPath = "\\$machineName\Storeman\Office\DBS\"

			if (-not (Test-Path -Path $targetPath))
			{
				Write_Log "Lane #${lane}: Target path '$targetPath' does not exist. Skipping." "Yellow"
				continue
			}

			Write_Log "Processing Lane #$lane at '$targetPath', please wait..." "Blue"

			try
			{
				# Delete_Files function is now expected to return an integer count
				$deletionCount = Delete_Files -Path $targetPath -SpecifiedFiles $fileExtensions -Exclusions @() -AsJob:$false

				if ($deletionCount -is [int])
				{
					$totalDeleted += $deletionCount
				}
				else
				{
					Write_Log "Lane #${lane}: Unexpected response from Delete_Files." "Red"
					$totalFailed++
				}
			}
			catch
			{
				Write_Log "Lane #${lane}: An error occurred while deleting files. Error: $_" "Red"
				$totalFailed++
			}
		}
		else
		{
			Write_Log "Lane #${lane}: Machine information not found. Skipping." "Yellow"
			continue
		}
	}

	# Summary of deletion results
	Write_Log "Deletion Summary for Store Number: $StoreNumber - Total Files Deleted: $totalDeleted, Total Failures: $totalFailed." "Blue"
	Write_Log "`r`n==================== Delete_DBS Function Completed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Close_Open_Transactions
# ---------------------------------------------------------------------------------------------------
# Description:
#   Closes open POS transactions on the lanes referenced by recent XE error files: scans the
#   store XE folder for S*.*.* files (max 30 days old), extracts the lane number from each
#   file name, marks the lane's open SAL_HDR transactions as CLOSE via SQL (file fallback),
#   and can send a restart-all-programs command over the mailslot after deployment.
#
# Parameters:
#   - StoreNumber  [string]: store to process (required)
#
# Variables:
#   - $script:FunctionResults: SqlModuleName and the lane maps (input)
#   - $OfficePath: Storeman Office root (XE folder)
#
# Writes:
#   - lane databases (SAL_HDR F1067 = CLOSE) and Scripts_by_Alex_C.T\Closed_Transactions_LOG.txt
# ===================================================================================================

function Close_Open_Transactions
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)

	# ==================== Start and validate XE path ====================
	Write_Log "`r`n==================== Starting Close_Open_Transactions ====================`r`n" "blue"

	$XEFolderPath = "$OfficePath\XE${StoreNumber}901"
	if (-not (Test-Path $XEFolderPath))
	{
		Write_Log "XE folder not found: $XEFolderPath" "red"
		return
	}

	# ==================== Validate SQL module availability ====================
	# Invoke-Sqlcmd for the direct close over the lane's protocol; without a SQL module the
	# close goes out as a Close_Transaction.sqi file drop instead
	$invokeSql = $null
	$supportsConnStr = $false
	$SqlModule = $script:FunctionResults['SqlModuleName']
	if ($SqlModule -and $SqlModule -ne "None")
	{
		try
		{
			Import-Module $SqlModule -ErrorAction Stop
			$invokeSql = Get-Command Invoke-Sqlcmd -ErrorAction Stop
			$supportsConnStr = $invokeSql.Parameters.ContainsKey('ConnectionString')
		}
		catch { $invokeSql = $null }
	}
	if (-not $invokeSql) { Write_Log "No SQL module available - transactions are closed through a file drop to the lane." "yellow" }

	# Close over the lane's SQL protocol (Named Pipes / TCP); $true when it worked
	$closeViaProtocol = {
		param ([string]$Lane, [string]$Protocol, $LaneInfo, [string]$Sql, [string]$What)
		if (-not $invokeSql -or $Protocol -notin @('Named Pipes', 'TCP') -or -not $LaneInfo) { return $false }
		$connStr = if ($Protocol -eq 'Named Pipes') { $LaneInfo['NamedPipesConnStr'] } else { $LaneInfo['TcpConnStr'] }
		try
		{
			if ($supportsConnStr) { $null = & $invokeSql -ConnectionString $connStr -Query $Sql -QueryTimeout 30 -ErrorAction Stop }
			else { $null = & $invokeSql -ServerInstance $LaneInfo['TcpServer'] -Database $LaneInfo['DBName'] -Query $Sql -QueryTimeout 30 -ErrorAction Stop }
			Write_Log "Closed $What via SQL protocol ($Protocol) on lane $Lane." "green"
			return $true
		}
		catch
		{
			Write_Log "$Protocol failed for lane $Lane." "yellow"
			return $false
		}
	}

	# Close through a Close_Transaction.sqi drop in the lane's XF folder; $true when written
	$closeViaFile = {
		param ([string]$Lane, [string]$Sqi)
		$laneDirectory = "$OfficePath\XF${StoreNumber}${Lane}"
		if (-not (Test-Path -LiteralPath $laneDirectory))
		{
			Write_Log "Lane directory $laneDirectory not found (file fallback)." "yellow"
			return $false
		}
		try
		{
			Write_Sms_File -Path (Join-Path $laneDirectory 'Close_Transaction.sqi') -Content $Sqi
			Write_Log "Wrote Close_Transaction.sqi for lane $Lane (file fallback)." "yellow"
			return $true
		}
		catch
		{
			Write_Log "Could not write Close_Transaction.sqi for lane ${Lane}: $($_.Exception.Message)" "red"
			return $false
		}
	}

	# SQL string for closing all open transactions (manual fallback)
	$CloseTransactionManual = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE')"

	# Paths for logs (keep a persistent transaction log)
	$LogFolderPath = "$BasePath\Scripts_by_Alex_C.T"
	$LogFilePath = Join-Path -Path $LogFolderPath -ChildPath "Closed_Transactions_LOG.txt"
	if (-not (Test-Path $LogFolderPath))
	{
		try { New-Item -Path $LogFolderPath -ItemType Directory -Force | Out-Null }
		catch { Write_Log "Failed to create log directory '$LogFolderPath'. Error: $_" "red"; return }
	}

	$MatchedTransactions = $false

	try
	{
		$currentTime = Get-Date
		# Only process XE files less than 30 days old and starting with S (SMSStart)
		$files = @(Get-ChildItem -Path $XEFolderPath -Filter "S*.???" -ErrorAction SilentlyContinue | Where-Object {
			($currentTime - $_.LastWriteTime).TotalDays -le 30
		})

		# ==================== PROCESS EACH XE ERROR FILE FOUND ====================
		if ($files -and $files.Count -gt 0)
		{
			foreach ($file in $files)
			{
				try
				{
					# Attempt to extract the lane number from the filename
					if ($file.Name -match '^S.*\.(\d{3})$') { $LaneNumber = $Matches[1] }
					else { continue }

					# Read contents for parsing
					$content = Get-Content -Path $file.FullName
					$fromLine = $content | Where-Object { $_ -like 'From:*' }
					$subjectLine = $content | Where-Object { $_ -like 'Subject:*' }
					$msgLine = $content | Where-Object { $_ -like 'MSG:*' }
					$lastRecordedStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' }

					# Parse out store/lane numbers from the "From:" line
					if ($fromLine -match 'From:\s*(\d{3})(\d{3})')
					{
						$fileStoreNumber = $Matches[1]
						$fileLaneNumber = $Matches[2]
						if ($fileStoreNumber -eq $StoreNumber -and $fileLaneNumber -eq $LaneNumber)
						{
							# Get subject/MSG and check if it matches open transaction/health check
							if ($subjectLine -match 'Subject:\s*(.*)')
							{
								$subject = $Matches[1].Trim()
								if ($subject -eq 'Health' -and $msgLine -match 'MSG:\s*(.*)' -and $Matches[1].Trim() -eq 'This application is not running.')
								{
									# Get transaction number from status line
									if ($lastRecordedStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
									{
										$transactionNumber = $Matches[1]
										$laneKey = $LaneNumber.PadLeft(3, '0')
										$protocolType = if ($script:LaneProtocols -and $script:LaneProtocols.ContainsKey($laneKey)) { $script:LaneProtocols[$laneKey] }
										else { "File" }
										if ($protocolType -eq 'Offline')
										{
											Write_Log "Lane $LaneNumber is offline. Skipping transaction close." "yellow"
											continue
										}

										# Get full DB connection info for this lane
										$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber
										$closeSQL = "UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1032 = $transactionNumber"
										$protocolWorked = & $closeViaProtocol $LaneNumber $protocolType $laneInfo $closeSQL "transaction $transactionNumber"

										# ----------- Fallback: File-based SQI if protocol failed -----------
										if (-not $protocolWorked)
										{
											$CloseTransactionAuto = "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1032 = $transactionNumber)"
											$protocolWorked = & $closeViaFile $LaneNumber $CloseTransactionAuto
										}
										if (-not $protocolWorked)
										{
											# Nothing reached the lane: keep the XE report, do not log a close or restart
											Write_Log "Transaction $transactionNumber on lane $LaneNumber NOT closed (no SQL access and no XF folder) - XE file kept." "red"
											continue
										}

										# Log the close action
										$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Closed transaction $transactionNumber on lane $LaneNumber"
										Add-Content -Path $LogFilePath -Value $logMessage

										# Remove processed XE file
										Remove-Item -Path $file.FullName -Force

										Write_Log "Processed file $($file.Name) for lane $LaneNumber and closed transaction $transactionNumber" "green"
										$MatchedTransactions = $true

										# ----------- Optionally restart all programs on the lane -----------
										Start-Sleep -Seconds 3
										$laneMap = $script:FunctionResults['LaneNumToMachineName']
										if ($laneMap)
										{
											$machineName = $laneMap[$LaneNumber]
											if ($machineName)
											{
												$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
												$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
												$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
												if ($result) { Write_Log "Restart command sent to Machine $machineName (Store $StoreNumber, Lane $LaneNumber) after deployment." "green" }
												else { Write_Log "Failed to send restart command to Machine $machineName (Store $StoreNumber, Lane $LaneNumber)." "red" }
											}
										}
									}
									else
									{
										Write_Log "Could not extract transaction number from Last recorded status in file $($file.Name)" "red"
									}
								}
							}
						}
					}
				}
				catch
				{
					Write_Log "Error processing file $($file.Name): $_" "red"
				}
			}
		}

		# ==================== If no matching files: Manual prompt for lane ====================
		if (-not $MatchedTransactions)
		{
			Write_Log "No files or no matching transactions found. Prompting for lane number." "yellow"
			$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
			if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
			{
				Write_Log "Lane selection cancelled or returned no selection." "yellow"
				Write_Log "`r`n==================== Close_Open_Transactions Function Completed ====================" "blue"
				return
			}

			foreach ($LaneNumber in $selection.Lanes)
			{
				$LaneNumber = if ($LaneNumber -is [pscustomobject] -and $LaneNumber.LaneNumber) { $LaneNumber.LaneNumber }
				else { $LaneNumber }
				$LaneNumber = $LaneNumber.PadLeft(3, '0')
				$laneKey = $LaneNumber
				$protocolType = if ($script:LaneProtocols -and $script:LaneProtocols.ContainsKey($laneKey)) { $script:LaneProtocols[$laneKey] }
				else { "File" }
				if ($protocolType -eq 'Offline')
				{
					Write_Log "Lane $LaneNumber is offline. Skipping." "yellow"
					continue
				}
				$laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber
				$closeSQL = "UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE'"
				$protocolWorked = & $closeViaProtocol $LaneNumber $protocolType $laneInfo $closeSQL "all open transactions"

				if (-not $protocolWorked) { $protocolWorked = & $closeViaFile $LaneNumber $CloseTransactionManual }
				if (-not $protocolWorked)
				{
					Write_Log "Lane ${LaneNumber}: open transactions NOT closed (no SQL access and no XF folder)." "red"
					continue
				}
				$logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - User deployed Close_Transaction to lane $LaneNumber"
				Add-Content -Path $LogFilePath -Value $logMessage
				$laneMap = $script:FunctionResults['LaneNumToMachineName']
				if ($laneMap)
				{
					$machineName = $laneMap[$LaneNumber]
					if ($machineName)
					{
						$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${LaneNumber}"
						$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."
						$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
						if ($result) { Write_Log "Restart All Programs sent to Machine $machineName (Store $StoreNumber, Lane $LaneNumber) after user deployment." "green" }
						else { Write_Log "Failed to send restart command to Machine $machineName (Store $StoreNumber, Lane $LaneNumber)." "red" }
					}
				}
				Write_Log "Prompt deployment process completed." "yellow"
			}
			# Clean up the XE folder, except FATALs
			Get-ChildItem -Path $XEFolderPath -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*FATAL*" } | Remove-Item -Force -ErrorAction SilentlyContinue
		}
	}
	catch
	{
		Write_Log "An error occurred during monitoring: $_" "red"
	}
	Write_Log "No further matching files were found after processing." "yellow"
	Write_Log "`r`n==================== Close_Open_Transactions Function Completed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Refresh_PIN_Pad_Files
# ---------------------------------------------------------------------------------------------------
# Description:
#   Refreshes the EMV PIN pad configuration files (PreferredAIDs.ini, EMVCAPKey.ini,
#   EMVAID.ini) on selected lanes: each target file is touched (LastWriteTime updated) so
#   the lane sees it as modified and reloads it.
#
# Parameters:
#   - StoreNumber  [string]: store to process, 3 digits (required)
#
# Variables:
#   - $script:FunctionResults['LaneNumToMachineName']: lane map (input)
#   - $OfficePath: Storeman Office root (XF lane folders)
#
# Notes:
#   - Retrieve_Nodes must have been run first.
# ===================================================================================================

function Refresh_PIN_Pad_Files
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber
	)

	Write_Log "`r`n==================== Starting Refresh_PIN_Pad_Files Function ====================`r`n" "blue"

	# ============================= Validate Environment =============================
	# Validate OfficePath base directory
	if (-not (Test-Path $OfficePath))
	{
		Write_Log "XF Base Path not found: $OfficePath" "yellow"
		return
	}

	# Ensure required functions are loaded
	foreach ($func in @('Show_Node_Selection_Form', 'Write_Log', 'Retrieve_Nodes'))
	{
		if (-not (Get-Command -Name $func -ErrorAction SilentlyContinue))
		{
			Write-Error "Function '$func' is not available. Please ensure it is loaded."
			return
		}
	}

	# Ensure lane mapping is available from Retrieve_Nodes
	if (-not ($script:FunctionResults.ContainsKey('LaneNumToMachineName')))
	{
		Write_Log "No lane information found. Please ensure Retrieve_Nodes has been executed." "red"
		return
	}
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']

	# ============================= Lane Selection UI =============================
	# Use the node selector form to allow the user to choose which lanes to refresh
	$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if ($null -eq $selection)
	{
		Write_Log "Operation canceled by user." "yellow"
		return
	}
	$Lanes = $selection.Lanes
	if (-not $Lanes -or $Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for processing." "yellow"
		return
	}

	# ============================= File Types =============================
	# List of PIN pad config files that should be "touched" (refreshed)
	$fileExtensions = @("PreferredAIDs.ini", "EMVCAPKey.ini", "EMVAID.ini")
	Write_Log "Starting refresh of file types: $($fileExtensions -join ', ') for Store Number: $StoreNumber." "green"

	# ============================= Main Processing Loop =============================
	$totalRefreshed = 0
	$totalFailed = 0
	$refreshedLanes = New-Object System.Collections.Generic.List[string]

	foreach ($laneObj in $Lanes)
	{
		# Accept both plain lane number string or PSCustomObject with .LaneNumber property
		$laneNum = if ($laneObj -is [pscustomobject] -and $laneObj.LaneNumber) { $laneObj.LaneNumber }
		else { $laneObj }
		$laneNum = $laneNum.PadLeft(3, '0')

		# Ensure this lane exists in LaneNumToMachineName
		if ($LaneNumToMachineName.ContainsKey($laneNum))
		{
			$machineName = $LaneNumToMachineName[$laneNum]

			if ([string]::IsNullOrWhiteSpace($machineName) -or $machineName -eq "Unknown")
			{
				Write_Log "Lane #${laneNum}: Machine name is invalid or unknown. Skipping refresh." "yellow"
				continue
			}

			# Build UNC path to lane's EMV config directory
			$targetPath = "\\$machineName\Storeman\XchDev\EMVConfig\"

			if (-not (Test-Path -Path $targetPath))
			{
				Write_Log "Lane #${laneNum}: Target path '$targetPath' does not exist. Skipping." "yellow"
				continue
			}

			Write_Log "Processing Lane #$laneNum ($machineName) at '$targetPath'." "blue"

			foreach ($file in $fileExtensions)
			{
				$filePath = Join-Path -Path $targetPath -ChildPath $file
				if (Test-Path -Path $filePath)
				{
					try
					{
						# "Touch" the file (set LastWriteTime to current time)
						(Get-Item -Path $filePath).LastWriteTime = Get-Date
						Write_Log "Lane #${laneNum}: Refreshed file '$filePath'." "green"
						$totalRefreshed++
						if (-not $refreshedLanes.Contains($laneNum)) { $refreshedLanes.Add($laneNum) }
					}
					catch
					{
						Write_Log "Lane #${laneNum}: Failed to refresh file '$filePath'. Error: $_" "red"
						$totalFailed++
					}
				}
				else
				{
					Write_Log "Lane #${laneNum}: File '$filePath' does not exist. Skipping." "yellow"
				}
			}
		}
		else
		{
			Write_Log "Lane #${laneNum}: Machine information not found in LaneNumToMachineName. Skipping." "yellow"
			continue
		}
	}
	Write_Log "Refresh Summary for Store Number: $StoreNumber - Total Files Refreshed: $totalRefreshed, Total Failures: $totalFailed." "green"
	# ==== Prompt to Restart All Programs on successful lanes ====
	if ($totalRefreshed -gt 0)
	{
		# ($refreshedLanes is collected in the loop above - lanes where at least one file was touched)
		if ($refreshedLanes.Count -gt 0)
		{
			Add-Type -AssemblyName System.Windows.Forms
			$laneListStr = ($refreshedLanes | Sort-Object) -join ", "
			$msg = "Do you want to send 'Restart All Programs' to the following lanes?`nStore $StoreNumber Lanes: $laneListStr"
			$result = [System.Windows.Forms.MessageBox]::Show($msg, "Restart All Programs?", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
			if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				Write_Log "User chose to send 'Restart All Programs' to lanes: $laneListStr" "cyan"
				Send_Restart_All_Programs -StoreNumber $StoreNumber -LaneNumbers $refreshedLanes
			}
			else
			{
				Write_Log "User cancelled 'Restart All Programs' action." "yellow"
			}
		}
	}
	Write_Log "`r`n==================== Refresh_PIN_Pad_Files Function Completed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Send_Restart_All_Programs
# ---------------------------------------------------------------------------------------------------
# Description:
#   Sends a restart-all-programs command to the selected lanes of a store: uses the cached
#   lane-to-machine map (Retrieve_Nodes when it is missing), picks lanes (or uses the passed
#   list) and sends "@exec(RESTART_ALL=PROGRAMS)." to each lane's SMSStart mailslot.
#
# Parameters:
#   - StoreNumber  [string]: store to process, 3 digits (required)
#   - LaneNumbers  [array]: lanes to target (default: selection form)
#   - Silent       [switch]: suppress Write_Log output
#
# Variables:
#   - $script:FunctionResults: LaneNumToMachineName (input via Retrieve_Nodes)
# ===================================================================================================

function Send_Restart_All_Programs
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[array]$LaneNumbers,
		[Parameter(Mandatory = $false)]
		[switch]$Silent
	)

	if (-not $Silent)
	{
		Write_Log "`r`n==================== Starting Send_Restart_All_Programs Function ====================`r`n" "blue"
	}

	# Lane map from the startup Retrieve_Nodes run; only re-run it when the map is missing
	$laneMap = $script:FunctionResults['LaneNumToMachineName']
	if (-not $laneMap -or $laneMap.Count -eq 0)
	{
		$nodes = Retrieve_Nodes -StoreNumber $StoreNumber
		if ($nodes) { $laneMap = $nodes.LaneNumToMachineName }
	}
	if (-not $laneMap)
	{
		Write_Log "Failed to retrieve node information for store $StoreNumber." "red"
		return
	}

	# Lane selection: use passed-in list or picker
	$lanes =
	if ($LaneNumbers -and $LaneNumbers.Count -gt 0)
	{
		$LaneNumbers | ForEach-Object {
			if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
			elseif ($_ -is [int]) { "{0:D3}" -f $_ }
			elseif ($_.Length -lt 3) { $_.PadLeft(3, '0') }
			else { $_ }
		}
	}
	else
	{
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
		if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
		{
			Write_Log "No lanes selected or selection cancelled. Exiting." "yellow"
			return
		}
		$selection.Lanes | ForEach-Object {
			if ($_ -is [pscustomobject] -and $_.LaneNumber) { $_.LaneNumber }
			elseif ($_ -is [int]) { "{0:D3}" -f $_ }
			elseif ($_.Length -lt 3) { $_.PadLeft(3, '0') }
			else { $_ }
		}
	}

	foreach ($lane in $lanes)
	{
		$machineName = $laneMap[$lane]
		if (-not $machineName)
		{
			Write_Log "No machine found for lane $lane. Skipping." "yellow"
			continue
		}

		$mailslotAddress = "\\$machineName\Mailslot\SMSStart_${StoreNumber}${lane}"
		$commandMessage = "@exec(RESTART_ALL=PROGRAMS)."

		try
		{
			$result = [MailslotSender]::SendMailslotCommand($mailslotAddress, $commandMessage)
			if ($result)
			{
				if (-not $Silent)
				{
					Write_Log "Command sent successfully to Machine $machineName (Store $StoreNumber, Lane $lane)." "green"
				}
			}
			else
			{
				if (-not $Silent)
				{
					Write_Log "Failed to send command to Machine $machineName (Store $StoreNumber, Lane $lane)." "red"
				}
			}
		}
		catch
		{
			if (-not $Silent)
			{
				Write_Log "Exception sending to $machineName (Lane $lane): $_" "red"
			}
		}
	}

	if (-not $Silent)
	{
		Write_Log "`r`n==================== Send_Restart_All_Programs Function Completed ====================" "blue"
	}
}

# ===================================================================================================
#                              FUNCTION: INI_Editor
# ---------------------------------------------------------------------------------------------------
# Description:
#   GUI INI editor for files under \storeman\<RelativeIniPath>: pick Server/Lane as source,
#   load the INI, select a section, edit all keys (add/edit/delete/reorder) and save. After
#   saving, the INI can be deployed to other lanes either as a merge of the edited section
#   or as a full file replacement. Copy stays disabled until a Save.
#
# Parameters:
#   - RelativeIniPath    [string]: INI path under the storeman root (default office\Setup.ini)
#   - PredefinedIniPaths [string[]]: entries for the source combo
#
# Variables:
#   - $script:FunctionResults: LaneMachines / LaneNumToMachineName (known lanes)
#   - $script:StoremanRoots: candidate local server storeman roots
#
# Writes:
#   - the edited INI (save) and the deployed copy/merge on the selected lanes
# ===================================================================================================

function INI_Editor
{
	param (
		[string]$RelativeIniPath = 'office\Setup.ini',
		[string[]]$PredefinedIniPaths = @('office\Setup.ini', 'office\Setting.ini', 'office\System.ini', 'XchDev\ApiVerifoneMX.ini')
	)

	Write_Log "`r`n==================== Starting INI_Editor ====================`r`n" "blue"

	# ==============================================================================================
	#                                    HELPERS: paths, parse, write
	# ==============================================================================================

	# Copy of an INI before it is overwritten (save / merge / replace on a lane), so a bad edit can
	# be put back: <BackupRoot>\INI_Editor\<yyyyMMdd>\<HHmmss>_<host>_<file>
	$backupIni = {
		param ([string]$Path)
		try
		{
			if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
			$hostPart = if ($Path -match '^\\\\([^\\]+)\\') { $matches[1] } else { $env:COMPUTERNAME }
			$dir = Join-Path $script:BackupRoot ("INI_Editor\" + (Get-Date -Format 'yyyyMMdd'))
			if (-not (Test-Path -LiteralPath $dir)) { $null = New-Item -Path $dir -ItemType Directory -Force }
			Copy-Item -LiteralPath $Path -Destination (Join-Path $dir ("{0}_{1}_{2}" -f (Get-Date -Format 'HHmmss'), $hostPart, (Split-Path -Leaf $Path))) -Force -ErrorAction Stop
		}
		catch { Write_Log "Could not back up ${Path} before writing: $($_.Exception.Message)" "yellow" }
	}

	$getKnownLanes = {
		$known = @()
		if ($script:FunctionResults.ContainsKey('LaneMachines') -and $script:FunctionResults['LaneMachines'])
		{ $known = $script:FunctionResults['LaneMachines'].Values | Where-Object { $_ } | Select-Object -Unique }
		elseif ($script:FunctionResults.ContainsKey('LaneNumToMachineName') -and $script:FunctionResults['LaneNumToMachineName'])
		{ $known = $script:FunctionResults['LaneNumToMachineName'].Values | Where-Object { $_ } | Select-Object -Unique }
		return, ($known | Sort-Object)
	}

	$getLaneRoot = {
		param ([string]$ComputerName)
		$s1 = "\\$ComputerName\Storeman"; if (Test-Path -LiteralPath $s1) { return $s1 }
		$s2 = "\\$ComputerName\c$\storeman"; if (Test-Path -LiteralPath $s2) { return $s2 }
		return $null
	}

	$getServerRoot = {
		foreach ($root in @($script:StoremanRoots))
		{
			if (Test-Path -LiteralPath $root) { return $root }
		}
		return "$($env:SystemDrive)\storeman"
	}

	$buildFullPath = {
		param ([string]$Root,
			[string]$Rel)
		$r = ($Rel -as [string]).Trim().TrimStart('\', '/')
		return ([System.IO.Path]::Combine($Root, $r))
	}

	# ---------- INI reader (keeps order of sections and keys in memory) ----------
	$readIni = {
		param ([string]$Path)
		if (-not (Test-Path -LiteralPath $Path)) { return $null }

		# ANSI (Windows-1252), as Storeman writes them: reading/writing ASCII turned every accented
		# character (store names, receipt lines) into "?" on save and on every merged lane
		$raw = @(Get-Content -LiteralPath $Path -Encoding Default)
		$sections = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]'
		$current = $null

		foreach ($line in $raw)
		{
			$t = $line.Trim()
			if ($t -match '^\s*[;#]') { continue }
			if ($t -match '^\[(?<sec>[^\]]+)\]\s*$')
			{
				$sec = $matches['sec']
				if (-not $sections.ContainsKey($sec))
				{
					$od = New-Object System.Collections.Specialized.OrderedDictionary
					$sections.Add($sec, $od)
				}
				$current = $sec
				continue
			}
			if ($current -and $t -match '^(?<k>[^=]+?)\s*=\s*(?<v>.*)$')
			{
				$k = $matches['k'].Trim()
				$v = $matches['v']
				if (-not $sections[$current].Contains($k)) { $sections[$current].Add($k, $v) }
				else { $sections[$current][$k] = $v }
			}
		}
		return @{ RawLines = $raw; Sections = $sections }
	}

	# ---------- INI section writer ----------
	# Merge (default): update keys in place, append new ones, leave other keys alone (deploys).
	# -ReplaceKeys: the section's keys become exactly $NewKeyValues, in that order (Save: keys
	# deleted or reordered in the grid used to stay as they were in the file).
	# Comment and blank lines stay where they are; no extra blank line is added.
	$writeIniSection = {
		param (
			[string[]]$RawLines,
			[string]$SectionName,
			[System.Collections.Specialized.OrderedDictionary]$NewKeyValues,
			[bool]$ReplaceKeys = $false
		)
		$lines = New-Object System.Collections.Generic.List[string]; $lines.AddRange($RawLines)
		$start = -1; $end = $lines.Count

		for ($i = 0; $i -lt $lines.Count; $i++)
		{
			if ($lines[$i] -match "^\s*\[$([regex]::Escape($SectionName))\]\s*$") { $start = $i; break }
		}

		if ($start -lt 0)
		{
			if ($lines.Count -gt 0 -and $lines[$lines.Count - 1].Trim() -ne '') { $lines.Add('') }
			$lines.Add("[$SectionName]")
			foreach ($k in $NewKeyValues.Keys) { $lines.Add("$k=$($NewKeyValues[$k])") }
			return, $lines.ToArray()
		}

		for ($i = $start + 1; $i -lt $lines.Count; $i++)
		{
			if ($lines[$i] -match '^\s*\[.+\]\s*$') { $end = $i; break }
		}

		if ($ReplaceKeys)
		{
			# drop every key line of the section, then write the new keys where the first one was
			$firstKey = -1
			for ($i = $end - 1; $i -gt $start; $i--)
			{
				$L = $lines[$i]
				if ($L.Trim() -notmatch '^[;#]' -and $L -match '^[^=]+=')
				{
					$lines.RemoveAt($i)
					$end--
					$firstKey = $i
				}
			}
			if ($firstKey -lt 0)
			{
				$firstKey = $start + 1
				for ($i = $end - 1; $i -gt $start; $i--) { if ($lines[$i].Trim() -ne '') { $firstKey = $i + 1; break } }
			}
			$offset = 0
			foreach ($k in $NewKeyValues.Keys) { $lines.Insert($firstKey + $offset, "$k=$($NewKeyValues[$k])"); $offset++ }
			return, $lines.ToArray()
		}

		# find last nonblank/noncomment inside section
		$lastContent = $start
		for ($i = [Math]::Min($end, $lines.Count) - 1; $i -gt $start; $i--)
		{
			$t = $lines[$i].Trim()
			if ($t -ne '' -and $t -notmatch '^[;#]') { $lastContent = $i; break }
		}
		$insertAt = $lastContent + 1

		# update existing keys
		$updated = @{ }
		for ($i = $start + 1; $i -lt $end; $i++)
		{
			$L = $lines[$i]
			if ($L.Trim() -match '^[;#]') { continue }
			if ($L -match '^(?<k>[^=]+?)\s*=\s*(?<v>.*)$')
			{
				$currKey = $matches['k'].Trim()
				$targetKey = $null
				foreach ($nk in $NewKeyValues.Keys) { if ($nk -ieq $currKey) { $targetKey = $nk; break } }
				if ($targetKey) { $lines[$i] = "$currKey=$($NewKeyValues[$targetKey])"; $updated[$targetKey] = $true }
			}
		}

		# add new keys (preserve order from $NewKeyValues)
		$pending = @()
		foreach ($k in $NewKeyValues.Keys) { if (-not $updated.ContainsKey($k)) { $pending += "$k=$($NewKeyValues[$k])" } }
		if ($pending.Count -gt 0)
		{
			$offset = 0
			foreach ($nl in $pending) { $lines.Insert($insertAt + $offset, $nl); $offset++ }
		}
		return, $lines.ToArray()
	}

	$lanesToMachines = {
		param ($LaneNums)
		$out = @()
		$map = $script:FunctionResults['LaneNumToMachineName']
		$known = & $getKnownLanes
		foreach ($ln in $LaneNums)
		{
			$pad = $null; try { $pad = "{0:D3}" -f ([int]$ln) }
			catch { $pad = "$ln" }
			$m = $null
			if ($map)
			{
				if ($map.ContainsKey($pad)) { $m = $map[$pad] }
				elseif ($map.ContainsKey([int]$ln)) { $m = $map[[int]$ln] }
				elseif ($map.ContainsKey("$([int]$ln)")) { $m = $map["$([int]$ln)"] }
			}
			if (-not $m)
			{
				$guess = $known | Where-Object { "$_" -like "*$pad" }
				if ($guess) { $m = $guess[0] }
			}
			if ($m) { $out += $m }
		}
		return, (@($out | Where-Object { $_ } | Select-Object -Unique))
	}

	# ==============================================================================================
	#                                           UI LAYOUT
	# ==============================================================================================

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	[void][System.Windows.Forms.Application]::EnableVisualStyles()

	$frm = New-Object System.Windows.Forms.Form
	$frm.Text = "INI_Editor  -  dynamic section/key editor"
	$frm.Size = New-Object System.Drawing.Size(820, 600) # keep size modest; we'll give the grid more room
	$frm.StartPosition = 'CenterScreen'
	$frm.FormBorderStyle = 'FixedDialog'
	$frm.MaximizeBox = $false
	$frm.MinimizeBox = $false
	$frm.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 251)

	# Root: 1 column; make grid take the most space (Percent=100 row at the end)
	$root = New-Object System.Windows.Forms.TableLayoutPanel
	$root.Dock = 'Fill'
	$root.Padding = New-Object System.Windows.Forms.Padding(8)
	$root.ColumnCount = 1
	$root.RowCount = 6
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # header
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # source group
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # INI row (flow)
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # Section row
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # GRID gets the rest
	$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # bottom buttons
	$frm.Controls.Add($root)

	# ---------- Header (small + tight margins so the grid gains space) ----------
	$hdr = New-Object System.Windows.Forms.Label
	$hdr.Text = "Pick Server/Lane, load an INI, choose/TYPE a section, edit keys, Save, then optionally deploy."
	$hdr.Font = New-Object System.Drawing.Font("Segoe UI", [float]9, [System.Drawing.FontStyle]::Bold)
	$hdr.AutoSize = $true
	$hdr.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 4)
	$root.Controls.Add($hdr, 0, 0)

	# ---------- Source group (compacted) ----------
	$grpSrc = New-Object System.Windows.Forms.GroupBox
	$grpSrc.Text = "Source"
	$grpSrc.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
	$grpSrc.Dock = 'Top'
	$root.Controls.Add($grpSrc, 0, 1)

	$srcGrid = New-Object System.Windows.Forms.TableLayoutPanel
	$srcGrid.Dock = 'Top'
	$srcGrid.ColumnCount = 2
	$srcGrid.RowCount = 2
	[void]$srcGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$srcGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$srcGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$srcGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$grpSrc.Controls.Add($srcGrid)

	$rbSrv = New-Object System.Windows.Forms.RadioButton
	$rbSrv.Text = "Server (this computer)"
	$rbSrv.Checked = $true
	$rbSrv.AutoSize = $true
	$rbSrv.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 2)
	$srcGrid.Controls.Add($rbSrv, 0, 0)
	$srcGrid.SetColumnSpan($rbSrv, 2)

	$rbLane = New-Object System.Windows.Forms.RadioButton
	$rbLane.Text = "Lane (source)"
	$rbLane.AutoSize = $true
	$rbLane.Margin = New-Object System.Windows.Forms.Padding(0, 4, 8, 2)
	$srcGrid.Controls.Add($rbLane, 0, 1)

	$laneInline = New-Object System.Windows.Forms.FlowLayoutPanel
	$laneInline.FlowDirection = 'LeftToRight'
	$laneInline.WrapContents = $false
	$laneInline.AutoSize = $true
	$laneInline.AutoSizeMode = 'GrowAndShrink'
	$laneInline.Dock = 'Fill'
	$laneInline.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
	$srcGrid.Controls.Add($laneInline, 1, 1)

	$lblLane = New-Object System.Windows.Forms.Label
	$lblLane.Text = "Select lane:"
	$lblLane.AutoSize = $true
	$lblLane.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 2)
	$laneInline.Controls.Add($lblLane)

	$cboLane = New-Object System.Windows.Forms.ComboBox
	$cboLane.DropDownStyle = 'DropDownList'
	$cboLane.Enabled = $false
	$cboLane.Width = 260
	$cboLane.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
	$laneInline.Controls.Add($cboLane)

	(& $getKnownLanes) | ForEach-Object { [void]$cboLane.Items.Add($_) }
	$rbLane.Add_CheckedChanged({
			$cboLane.Enabled = $rbLane.Checked
			if ($rbLane.Checked -and $cboLane.Items.Count -gt 0 -and -not $cboLane.SelectedItem) { $cboLane.SelectedIndex = 0 }
		})

	# ---------- INI path row (ONE ROW via FlowLayoutPanel; resizes combo so the button never wraps) ----------
	$iniFlow = New-Object System.Windows.Forms.FlowLayoutPanel
	$iniFlow.FlowDirection = 'LeftToRight'
	$iniFlow.WrapContents = $false
	$iniFlow.Dock = 'Top'
	$iniFlow.AutoSize = $false
	$iniFlow.Height = 38
	$iniFlow.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
	$root.Controls.Add($iniFlow, 0, 2)

	$lblRel = New-Object System.Windows.Forms.Label
	$lblRel.Text = "INI path under \storeman\"
	$lblRel.AutoSize = $true
	$lblRel.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)
	$iniFlow.Controls.Add($lblRel)

	$cboRel = New-Object System.Windows.Forms.ComboBox
	$cboRel.DropDownStyle = 'DropDown' # editable
	$cboRel.Margin = New-Object System.Windows.Forms.Padding(6, 3, 6, 3)
	$cboRel.Width = 520
	foreach ($p in ($PredefinedIniPaths | Select-Object -Unique)) { [void]$cboRel.Items.Add($p) }
	if ($RelativeIniPath -and -not ($cboRel.Items -contains $RelativeIniPath)) { [void]$cboRel.Items.Add($RelativeIniPath) }
	$cboRel.Text = $RelativeIniPath
	$iniFlow.Controls.Add($cboRel)

	$btnLoad = New-Object System.Windows.Forms.Button
	$btnLoad.Text = "Load INI"
	$btnLoad.Width = 100
	$btnLoad.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
	$iniFlow.Controls.Add($btnLoad)

	# Resize handler to keep all three controls on ONE line
	$iniFlow.Add_Resize({
			$total = $iniFlow.ClientSize.Width
			$wLabel = $lblRel.PreferredSize.Width
			$wBtn = $btnLoad.Width
			$pad = 24 # rough margins between controls
			$newW = [Math]::Max(220, $total - $wLabel - $wBtn - $pad)
			$cboRel.Width = $newW
		})

	# ---------- Section row (tight) ----------
	$rowSec = New-Object System.Windows.Forms.TableLayoutPanel
	$rowSec.Dock = 'Top'
	$rowSec.ColumnCount = 4
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 320)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
	[void]$rowSec.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$rowSec.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 2)
	$root.Controls.Add($rowSec, 0, 3)

	$lblSec = New-Object System.Windows.Forms.Label
	$lblSec.Text = "Section:"
	$lblSec.AutoSize = $true
	$lblSec.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 0)
	$rowSec.Controls.Add($lblSec, 0, 0)

	$cboSection = New-Object System.Windows.Forms.ComboBox
	$cboSection.DropDownStyle = 'DropDown'
	$cboSection.Width = 320
	$cboSection.Margin = New-Object System.Windows.Forms.Padding(0, 3, 6, 3)
	$rowSec.Controls.Add($cboSection, 1, 0)

	$lblPath = New-Object System.Windows.Forms.Label
	$lblPath.Text = ""
	$lblPath.AutoSize = $true
	$lblPath.ForeColor = [System.Drawing.Color]::DimGray
	$lblPath.Margin = New-Object System.Windows.Forms.Padding(12, 6, 6, 0)
	$rowSec.Controls.Add($lblPath, 2, 0)

	$btnReload = New-Object System.Windows.Forms.Button
	$btnReload.Text = "Reload"
	$btnReload.Width = 90
	$btnReload.Margin = New-Object System.Windows.Forms.Padding(6, 3, 6, 3)
	$rowSec.Controls.Add($btnReload, 3, 0)

	# ---------- Grid (make it take the space) ----------
	$grid = New-Object System.Windows.Forms.DataGridView
	$grid.Dock = 'Fill'
	$grid.AllowUserToAddRows = $false
	$grid.AllowUserToDeleteRows = $true
	$grid.AllowDrop = $true
	$grid.AutoSizeColumnsMode = 'Fill'
	$grid.RowHeadersVisible = $false
	$grid.SelectionMode = 'CellSelect'
	$grid.MultiSelect = $false
	$grid.BackgroundColor = [System.Drawing.Color]::White
	$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(242, 243, 245)
	$grid.EnableHeadersVisualStyles = $false
	$root.Controls.Add($grid, 0, 4)

	$dt = New-Object System.Data.DataTable
	[void]$dt.Columns.Add('Key', [string])
	[void]$dt.Columns.Add('Value', [string])
	$grid.DataSource = $dt
	$grid.EditMode = 'EditOnEnter'
	$grid.Columns['Key'].SortMode = 'NotSortable'
	$grid.Columns['Value'].SortMode = 'NotSortable'

	# ---------- Bottom buttons ----------
	$btnRow = New-Object System.Windows.Forms.FlowLayoutPanel
	$btnRow.FlowDirection = 'RightToLeft'
	$btnRow.Dock = 'Bottom'
	$btnRow.AutoSize = $true
	$btnRow.Padding = New-Object System.Windows.Forms.Padding(0, 6, 6, 6)
	$frm.Controls.Add($btnRow)

	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = "Close"
	$btnClose.Width = 90
	$btnRow.Controls.Add($btnClose)

	$btnCopy = New-Object System.Windows.Forms.Button
	$btnCopy.Text = "Copy to other lanes..."
	$btnCopy.Width = 170
	$btnCopy.Enabled = $false # stays off until a successful Save
	$btnRow.Controls.Add($btnCopy)

	$btnSave = New-Object System.Windows.Forms.Button
	$btnSave.Text = "Save to source"
	$btnSave.Width = 130
	$btnRow.Controls.Add($btnSave)

	$frm.AcceptButton = $btnSave
	$frm.CancelButton = $btnClose

	# ==============================================================================================
	#                                   CONTEXT MENU: Add/Delete
	# ==============================================================================================

	$ctx = New-Object System.Windows.Forms.ContextMenuStrip
	$miAdd = $ctx.Items.Add("Add key")
	$miDel = $ctx.Items.Add("Delete key")
	$grid.ContextMenuStrip = $ctx

	$script:_ctxRow = -1
	$grid.Add_CellMouseDown({
			param ($sender,
				$e)
			if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right)
			{
				if ($e.RowIndex -ge 0)
				{
					$grid.ClearSelection()
					$script:_ctxRow = $e.RowIndex
					$col = if ($e.ColumnIndex -ge 0) { $e.ColumnIndex }
					else { 0 }
					$grid.CurrentCell = $grid.Rows[$e.RowIndex].Cells[$col]
					$grid.Rows[$e.RowIndex].Selected = $true
				}
				else { $script:_ctxRow = -1 }
			}
		})

	# Add key at END; start editing; mark dirty (Copy disabled)
	$miAdd.Add_Click({
			$insertAt = $dt.Rows.Count
			$newRow = $dt.NewRow()
			$newRow['Key'] = ''
			$newRow['Value'] = ''
			[void]$dt.Rows.Add($newRow)

			$grid.Focus()
			try { $grid.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $insertAt) }
			catch { }
			$grid.CurrentCell = $grid.Rows[$insertAt].Cells[0]
			$grid.BeginEdit($true)

			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})

	$miDel.Add_Click({
			$rowIdx = -1
			if ($grid.CurrentCell) { $rowIdx = $grid.CurrentCell.RowIndex }
			if ($rowIdx -lt 0 -and $script:_ctxRow -ge 0) { $rowIdx = $script:_ctxRow }
			if ($rowIdx -lt 0 -or $rowIdx -ge $grid.Rows.Count) { return }

			$ans = [System.Windows.Forms.MessageBox]::Show(
				"Delete key '" + [string]$dt.Rows[$rowIdx]['Key'] + "'?",
				"Confirm delete", [System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question)
			if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

			$dt.Rows.RemoveAt($rowIdx)

			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})

	# Any edit => dirty => Copy disabled
	$grid.Add_CellValueChanged({ $state.IsDirty = $true; $btnCopy.Enabled = $false })
	$grid.Add_UserDeletedRow({ $state.IsDirty = $true; $btnCopy.Enabled = $false })

	# ==============================================================================================
	#                               DRAG & DROP ROW REORDERING
	# ==============================================================================================

	$script:_dragRow = -1
	$script:_dragging = $false
	$script:_dragStart = [System.Drawing.Point]::Empty

	$grid.Add_MouseDown({
			param ($s,
				$e)
			$script:_dragStart = [System.Drawing.Point]::new($e.X, $e.Y)
			$hit = $grid.HitTest($e.X, $e.Y)
			if ($hit.RowIndex -ge 0 -and $hit.RowIndex -lt $dt.Rows.Count) { $script:_dragRow = $hit.RowIndex }
			else { $script:_dragRow = -1 }
		})

	$grid.Add_MouseMove({
			param ($s,
				$e)
			if ($e.Button -band [System.Windows.Forms.MouseButtons]::Left -and $script:_dragRow -ge 0 -and -not $script:_dragging)
			{
				$dx = [Math]::Abs($e.X - $script:_dragStart.X)
				$dy = [Math]::Abs($e.Y - $script:_dragStart.Y)
				if ($dx -gt 4 -or $dy -gt 4)
				{
					$script:_dragging = $true
					$null = $grid.DoDragDrop("row-move", [System.Windows.Forms.DragDropEffects]::Move)
					$script:_dragging = $false
				}
			}
		})

	$grid.Add_DragOver({
			param ($s,
				$e)
			if ($script:_dragRow -lt 0) { $e.Effect = [System.Windows.Forms.DragDropEffects]::None; return }
			$pt = $grid.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
			$hit = $grid.HitTest($pt.X, $pt.Y)
			$target = if ($hit.RowIndex -lt 0) { $dt.Rows.Count - 1 }
			else { $hit.RowIndex }
			$e.Effect = if ($target -ne $script:_dragRow) { [System.Windows.Forms.DragDropEffects]::Move }
			else { [System.Windows.Forms.DragDropEffects]::None }
		})

	$grid.Add_DragDrop({
			param ($s,
				$e)
			if ($script:_dragRow -lt 0) { return }
			$source = $script:_dragRow
			$script:_dragRow = -1

			$pt = $grid.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
			$hit = $grid.HitTest($pt.X, $pt.Y)
			$target = $hit.RowIndex
			if ($target -lt 0) { $target = $dt.Rows.Count - 1 }
			if ($target -gt $dt.Rows.Count - 1) { $target = $dt.Rows.Count - 1 }
			if ($target -eq $source) { return }

			$keyVal = [string]$dt.Rows[$source]['Key']
			$valueVal = [string]$dt.Rows[$source]['Value']

			$dt.Rows.RemoveAt($source)
			if ($target -gt $source) { $target-- }

			$newRow = $dt.NewRow()
			$newRow['Key'] = $keyVal
			$newRow['Value'] = $valueVal
			if ($target -ge 0 -and $target -lt $dt.Rows.Count) { $dt.Rows.InsertAt($newRow, $target) }
			else { [void]$dt.Rows.Add($newRow); $target = $dt.Rows.Count - 1 }

			try { $grid.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $target) }
			catch { }
			$grid.CurrentCell = $grid.Rows[$target].Cells[0]
			$grid.Rows[$target].Selected = $true

			$state.IsDirty = $true
			$btnCopy.Enabled = $false
		})

	# ==============================================================================================
	#                                STATE + LOAD / SAVE / COPY
	# ==============================================================================================

	$state = @{
		IsDirty	       = $false
		SourceIsServer = $true
		SourceLane	   = $null
		FullPath	   = $null
		IniObj		   = $null
		DidSave	       = $false
		DidDeploy	   = $false
	}

	$loadSectionFromCombo = {
		$dt.Rows.Clear()
		if (-not $state.IniObj) { return }
		$secRequested = ($cboSection.Text -as [string]).Trim()
		if (-not $secRequested) { return }
		$secActual = $null
		foreach ($k in $state.IniObj.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
		if (-not $secActual) { return }
		$od = $state.IniObj.Sections[$secActual]
		foreach ($k in $od.Keys) { [void]$dt.Rows.Add($k, [string]$od[$k]) }
		$state.IsDirty = $false
		$btnCopy.Enabled = $false
	}

	$doLoad = {
		$rel = ($cboRel.Text -as [string]).Trim()
		if (-not $rel)
		{
			[System.Windows.Forms.MessageBox]::Show("Enter or select a relative INI path (e.g. office\Setup.ini).", "Missing path",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
			return
		}

		if ($rbSrv.Checked)
		{
			$state.SourceIsServer = $true
			$state.SourceLane = $null
			$rootSrv = & $getServerRoot
			$full = & $buildFullPath $rootSrv $rel
		}
		else
		{
			if (-not $cboLane.SelectedItem)
			{
				[System.Windows.Forms.MessageBox]::Show("Select a lane.", "Missing lane",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
				return
			}
			$state.SourceIsServer = $false
			$state.SourceLane = [string]$cboLane.SelectedItem
			$rootLane = & $getLaneRoot $state.SourceLane
			if (-not $rootLane)
			{
				[System.Windows.Forms.MessageBox]::Show("Lane storeman root not accessible (\\$($state.SourceLane)\Storeman or \\$($state.SourceLane)\C$\storeman).", "Path error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
				return
			}
			$full = & $buildFullPath $rootLane $rel
		}

		$state.FullPath = $full
		$lblPath.Text = $state.FullPath

		$ini = & $readIni $state.FullPath
		if (-not $ini)
		{
			Write_Log "[Info] INI not found, will be created on save: $($state.FullPath)" "yellow"
			$ini = @{
				RawLines = @("; Created by INI_Editor", "")
				Sections = (New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]')
			}
		}
		$state.IniObj = $ini

		$cboSection.Items.Clear()
		foreach ($secName in $state.IniObj.Sections.Keys) { [void]$cboSection.Items.Add($secName) }
		if ($cboSection.Items.Count -gt 0) { $cboSection.SelectedIndex = 0 }
		else { $cboSection.Text = "" }

		$loadSectionFromCombo.Invoke()
	}

	$btnLoad.Add_Click({ & $doLoad })
	$btnReload.Add_Click({ & $doLoad })
	$cboSection.Add_SelectedIndexChanged({ $loadSectionFromCombo.Invoke() })
	$cboSection.Add_TextChanged({ $loadSectionFromCombo.Invoke() })

	# ---------- Save ----------
	$btnSave.Add_Click({
			if (-not $state.FullPath)
			{
				[System.Windows.Forms.MessageBox]::Show("Load an INI first.", "No file",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			if (-not $state.IniObj) { return }

			$secRequested = ($cboSection.Text -as [string]).Trim()
			if (-not $secRequested)
			{
				[System.Windows.Forms.MessageBox]::Show("Enter or select a section name first.", "No section",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}

			$kv = New-Object System.Collections.Specialized.OrderedDictionary
			foreach ($row in $dt.Rows)
			{
				$k = [string]$row['Key']; $v = [string]$row['Value']
				if ([string]::IsNullOrWhiteSpace($k)) { continue }
				if ($kv.Contains($k)) { $kv.Remove($k) } # last one wins
				$kv.Add($k, $v)
			}

			$secActual = $null
			foreach ($k in $state.IniObj.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
			if (-not $secActual)
			{
				$secActual = $secRequested
				$od = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($k in $kv.Keys) { $od.Add($k, $kv[$k]) }
				$state.IniObj.Sections.Add($secActual, $od)
			}
			else
			{
				$od = $state.IniObj.Sections[$secActual]
				foreach ($k in @($od.Keys)) { $od.Remove($k) }
				foreach ($k in $kv.Keys) { $od.Add($k, $kv[$k]) }
			}

			$state.IniObj.RawLines = & $writeIniSection $state.IniObj.RawLines $secActual $kv $true

			try
			{
				$dir = Split-Path -Parent $state.FullPath
				if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
				& $backupIni $state.FullPath
				Set-Content -LiteralPath $state.FullPath -Value $state.IniObj.RawLines -Encoding Default -Force

				Write_Log "Saved [$secActual] to $($state.FullPath)" "green"
				[System.Windows.Forms.MessageBox]::Show("Saved section [$secActual] to:`r`n$($state.FullPath)", "Saved",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

				$state.DidSave = $true
				$state.IsDirty = $false
				$btnCopy.Enabled = $true
			}
			catch
			{
				Write_Log "Save failed for $($state.FullPath): $($_.Exception.Message)" "red"
				[System.Windows.Forms.MessageBox]::Show("Save failed:`r`n$($_.Exception.Message)", "Error",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
			}
		})

	# ---------- Copy (blocked if dirty or not saved) ----------
	$btnCopy.Add_Click({
			if ($state.IsDirty)
			{
				[System.Windows.Forms.MessageBox]::Show("Please Save first. Copy is disabled until the INI is saved.", "Save required",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}
			if (-not $state.FullPath -or -not (Test-Path -LiteralPath $state.FullPath))
			{
				[System.Windows.Forms.MessageBox]::Show("Nothing to copy. Save the INI first.", "No file",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}

			$secRequested = ($cboSection.Text -as [string]).Trim()

			# deploy dialog
			$dlg = New-Object System.Windows.Forms.Form
			$dlg.Text = "Deploy options"
			$dlg.StartPosition = 'CenterParent'
			$dlg.FormBorderStyle = 'FixedDialog'
			$dlg.MaximizeBox = $false
			$dlg.MinimizeBox = $false
			$dlg.ClientSize = New-Object System.Drawing.Size(460, 160)

			$lbl = New-Object System.Windows.Forms.Label
			$lbl.AutoSize = $true
			$lbl.MaximumSize = New-Object System.Drawing.Size(440, 0)
			$lbl.Location = New-Object System.Drawing.Point(10, 10)
			$lbl.Text = "How would you like to deploy to other lanes?`r`n`r`n" +
			"* Merge - update only section [$secRequested] keys on targets.`r`n" +
			"* Copy  - replace the entire INI file on targets."
			$dlg.Controls.Add($lbl)

			$btnMerge = New-Object System.Windows.Forms.Button
			$btnMerge.Text = "Merge"; $btnMerge.Size = New-Object System.Drawing.Size(90, 28)
			$btnMerge.Location = New-Object System.Drawing.Point(140, 110)
			$btnMerge.DialogResult = [System.Windows.Forms.DialogResult]::Yes

			$btnCopyAll = New-Object System.Windows.Forms.Button
			$btnCopyAll.Text = "Copy"; $btnCopyAll.Size = New-Object System.Drawing.Size(90, 28)
			$btnCopyAll.Location = New-Object System.Drawing.Point(240, 110)
			$btnCopyAll.DialogResult = [System.Windows.Forms.DialogResult]::No

			$btnCancel = New-Object System.Windows.Forms.Button
			$btnCancel.Text = "Cancel"; $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
			$btnCancel.Location = New-Object System.Drawing.Point(340, 110)
			$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

			$dlg.Controls.AddRange(@($btnMerge, $btnCopyAll, $btnCancel))
			$dlg.AcceptButton = $btnMerge
			$dlg.CancelButton = $btnCancel
			if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $dlg }

			$choice = $dlg.ShowDialog()
			if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
			$mergeOnly = ($choice -eq [System.Windows.Forms.DialogResult]::Yes)

			$excluded = @()
			if (-not $state.SourceIsServer -and $state.SourceLane) { $excluded = @($state.SourceLane) }

			$StoreNumber = $script:FunctionResults['StoreNumber']
			$sel = $null
			$sel = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane" -Title "Select destination lanes" -ExcludedNodes $excluded
			if (-not $sel) { Write_Log "Copy aborted: no lanes selected." "yellow"; return }

			$laneNums = @()
			if ($sel -is [System.Collections.IDictionary])
			{
				if ($sel.Contains('Lanes') -and $sel['Lanes']) { $laneNums = @($sel['Lanes']) }
				elseif ($sel.Contains('LaneNumbers') -and $sel['LaneNumbers']) { $laneNums = @($sel['LaneNumbers']) }
			}
			elseif ($sel -is [System.Collections.IEnumerable] -and -not ($sel -is [string])) { $laneNums = @($sel) }
			elseif ($sel -is [string]) { $laneNums = @($sel -split '[,\s]+' | Where-Object { $_ }) }
			$laneNums = @($laneNums | ForEach-Object {
					try { "{0:D3}" -f ([int]$_) }
					catch { "$_" }
				}) | Where-Object { $_ -match '^\d{3}$' }
			if (-not $laneNums -or $laneNums.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("No destination lanes selected.", "No selection",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}

			$destMachines = & $lanesToMachines $laneNums
			if (-not $state.SourceIsServer -and $state.SourceLane) { $destMachines = $destMachines | Where-Object { $_ -ne $state.SourceLane } }
			$destMachines = @($destMachines | Select-Object -Unique)
			if ($destMachines.Count -eq 0)
			{
				[System.Windows.Forms.MessageBox]::Show("No valid destination machines resolved.", "No destinations",
					[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
				return
			}

			$relPath = ($cboRel.Text -as [string]).Trim()
			$srcDir = Split-Path -Parent $state.FullPath
			$srcFile = Split-Path -Leaf $state.FullPath
			$ok = 0; $fail = 0

			if ($mergeOnly)
			{
				$kv = New-Object System.Collections.Specialized.OrderedDictionary
				foreach ($row in $dt.Rows)
				{
					$k = [string]$row['Key']; $v = [string]$row['Value']
					if ([string]::IsNullOrWhiteSpace($k)) { continue }
					if ($kv.Contains($k)) { $kv.Remove($k) }
					$kv.Add($k, $v)
				}

				foreach ($m in $destMachines)
				{
					$rootLane = & $getLaneRoot $m
					if (-not $rootLane) { Write_Log "[Dest Unreachable] $m has no \\Storeman nor \\C$\storeman." "red"; $fail++; continue }
					$dstFull = & $buildFullPath $rootLane $relPath
					$dstDir = Split-Path -Parent $dstFull
					try { if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null } }
					catch { Write_Log "[Create Failed] $dstDir : $($_.Exception.Message)" "red"; $fail++; continue }

					$destIni = & $readIni $dstFull
					if (-not $destIni)
					{
						$destIni = @{ RawLines = @("; Created by INI_Editor (merge)", ""); Sections = (New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Specialized.OrderedDictionary]') }
					}
					$secActual = $null
					foreach ($k in $destIni.Sections.Keys) { if ($k -ieq $secRequested) { $secActual = $k; break } }
					if (-not $secActual) { $secActual = $secRequested }

					$destIni.RawLines = & $writeIniSection $destIni.RawLines $secActual $kv
					try
					{
						& $backupIni $dstFull
						Set-Content -LiteralPath $dstFull -Value $destIni.RawLines -Encoding Default -Force
						Write_Log "Merged [$secActual] into $dstFull on $m" "green"; $ok++
					}
					catch
					{
						Write_Log "Merge failed on $m ($dstFull): $($_.Exception.Message)" "red"; $fail++
					}
				}
			}
			else
			{
				foreach ($m in $destMachines)
				{
					$rootLane = & $getLaneRoot $m
					if (-not $rootLane) { Write_Log "[Dest Unreachable] $m has no \\Storeman nor \\C$\storeman." "red"; $fail++; continue }
					$dstFull = & $buildFullPath $rootLane $relPath
					$dstDir = Split-Path -Parent $dstFull
					try { if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null } }
					catch { Write_Log "[Create Failed] $dstDir : $($_.Exception.Message)" "red"; $fail++; continue }

					& $backupIni (Join-Path $dstDir $srcFile)
					Write_Log "Replacing INI on $m ($dstDir\$srcFile)" "gray"
					$roboArgs = @($srcDir, $dstDir, $srcFile, '/COPY:DAT', '/R:2', '/W:2', '/NFL', '/NDL', '/NP', '/MT:8')
					$null = & robocopy @roboArgs
					$code = $LASTEXITCODE
					if ($code -ge 8) { Write_Log "Copy to $m FAILED (robocopy $code)" "red"; $fail++ }
					else { Write_Log "Copy to $m OK (robocopy $code)" "green"; $ok++ }
				}
			}
			if ($ok -gt 0) { $state.DidDeploy = $true } # <- at least one target updated OK
			[System.Windows.Forms.MessageBox]::Show("Deploy complete.`r`nOK: $ok   Fail: $fail", "Done",
				[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
		})

	# Activity line when the form closes (covers X, Alt+F4 and the Close button)
	$frm.Add_FormClosed({
			if (-not $state.DidSave -and -not $state.DidDeploy)
			{
				Write_Log "[No operations] Editor closed without saving or deploying." "gray"
			}
			else
			{
				# Optional: a tiny summary
				$acts = @()
				if ($state.DidSave) { $acts += "save" }
				if ($state.DidDeploy) { $acts += "deploy" }
				Write_Log "[Activity] Completed: $($acts -join ', ')." "gray"
			}
		})

	$btnClose.Add_Click({ $frm.Close() })

	# ==============================================================================================
	#                                         STARTUP
	# ==============================================================================================
	& $doLoad
	$iniFlow.PerformLayout() # ensure combo gets sized on first draw
	if ($Apply_Modern_Popup_Style) { & $Apply_Modern_Popup_Style $frm }
	[void]$frm.ShowDialog()

	Write_Log "`r`n==================== INI_Editor Completed ====================`r`n" "blue"
}

# ===================================================================================================
#                              FUNCTION: Remove_Archive_Bit
# ---------------------------------------------------------------------------------------------------
# Description:
#   Clears the file archive bit so the Storeman load process re-copies changed files. Run Now
#   does a one-time recursive clear using the lane map (or Ter_Load.sql parsing, local paths
#   only). Service mode installs an NSSM service with FileSystemWatchers that do a full
#   cleanup at startup and then clear the bits on each touched item as it changes; a
#   Ter_Load.sql change or a watcher overflow rebuilds the watchers with a full cleanup.
#   Includes an Uninstall Service button.
#
# Variables:
#   - $OfficePath: local Office root (input)
#   - $script:ScriptsFolder: folder for the generated service PS1
#   - $script:FunctionResults['StoreNumber']: store number for the logs
#
# Notes:
#   - Uses an already-downloaded nssm.exe; it never downloads NSSM itself.
# ===================================================================================================

function Remove_Archive_Bit
{
	[CmdletBinding()]
	param ()

	Write_Log "`r`n==================== Starting Remove_Archive_Bit Function ====================`r`n" "blue"

	# ---------------------------
	# NO-PROMPT "always sensing" poll interval (seconds)
	# ---------------------------
	$AutoPollSeconds = 3

	# ------------------------------------------------------------------------------------------------
	# Validate required context
	# ------------------------------------------------------------------------------------------------
	if (-not $OfficePath -or [string]::IsNullOrWhiteSpace($OfficePath))
	{
		Write_Log "ERROR: OfficePath is not set. Cannot continue." "red"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# --- Main context variables
	$storeNumber = [string]$script:FunctionResults['StoreNumber']
	$scriptFolder = $script:ScriptsFolder

	# ------------------------------------------------------------------------------------------------
	# Resolve OfficePath to LOCAL filesystem path when OfficePath is a local UNC share
	#   Example: \\localhost\storeman\office  ->  C:\Storeman\Office (if share "storeman" points to C:\Storeman)
	# ------------------------------------------------------------------------------------------------
	$OfficePathLocalResolved = $OfficePath
	try
	{
		$op = [string]$OfficePath
		if ($op -match '^(?i)\\\\([^\\]+)\\([^\\]+)(\\.*)?$')
		{
			$uncHost = [string]$matches[1]
			$share = [string]$matches[2]
			$rest = [string]$matches[3]

			$isLocalHost = $false
			if ($uncHost -match '^(?i)localhost$') { $isLocalHost = $true }
			elseif ($uncHost -eq $env:COMPUTERNAME) { $isLocalHost = $true }
			elseif ($uncHost -eq "$($env:COMPUTERNAME).$($env:USERDNSDOMAIN)") { $isLocalHost = $true }

			if ($isLocalHost)
			{
				$rest2 = ""
				if ($rest) { $rest2 = $rest.TrimStart('\') }

				$sharePath = $null
				try
				{
					$sm = Get-SmbShare -Name $share -ErrorAction Stop
					if ($sm -and $sm.Path) { $sharePath = [string]$sm.Path }
				}
				catch { }

				if (-not $sharePath)
				{
					try
					{
						$w = Get-CimInstance -ClassName Win32_Share -Filter ("Name='{0}'" -f $share.Replace("'", "''")) -ErrorAction Stop
						if ($w -and $w.Path) { $sharePath = [string]$w.Path }
					}
					catch { }
				}

				if (-not [string]::IsNullOrWhiteSpace($sharePath))
				{
					if (-not [string]::IsNullOrWhiteSpace($rest2))
					{
						$OfficePathLocalResolved = Join-Path $sharePath $rest2
					}
					else
					{
						$OfficePathLocalResolved = $sharePath
					}
				}
			}
		}
	}
	catch { }

	# IMPORTANT: build Ter_Load paths AFTER local resolution (service must not depend on UNC)
	$terFileRaw = Join-Path $OfficePath "Load\Ter_Load.sql"
	$terFileLocal = Join-Path $OfficePathLocalResolved "Load\Ter_Load.sql"
	$terFile = $terFileLocal
	if (-not (Test-Path -LiteralPath $terFileLocal) -and (Test-Path -LiteralPath $terFileRaw))
	{
		# Fallback only if local doesn't exist
		$terFile = $terFileRaw
	}

	# Optional: log what we're using (helps confirm fix)
	Write_Log "OfficePath (raw): $OfficePath" "yellow"
	Write_Log "OfficePath (local resolved): $OfficePathLocalResolved" "yellow"
	Write_Log "Ter_Load (local preferred): $terFileLocal" "yellow"
	Write_Log "Ter_Load (selected): $terFile" "yellow"

	# ------------------------------------------------------------------------------------------------
	# Service install paths
	# ------------------------------------------------------------------------------------------------
	$installRoot = Join-Path $scriptFolder "Remove_Archive_Bit_Service"
	$svcScripts = Join-Path $installRoot "Scripts"
	$svcLogsDir = Join-Path $installRoot "Logs"
	$svcPS1 = Join-Path $svcScripts "Remove_Archive_Bit_Service.ps1"
	$svcConfigPath = Join-Path $installRoot "Remove_Archive_Bit_Config.json"
	$serviceName = "Remove_Archive_Bit"

	foreach ($d in @($installRoot, $svcScripts, $svcLogsDir))
	{
		if (-not (Test-Path -LiteralPath $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
	}

	# ------------------------------------------------------------------------------------------------
	# NSSM: Use already-downloaded NSSM only (NO downloads here)
	#   We will accept:
	#     - $script:FunctionResults['NssmExe'] (if your startup Download_NSSM set it)
	#     - $script:NssmExe (if set globally)
	#     - Default global path: C:\Tecnica_Systems\Alex_C.T\Tools\NSSM\nssm.exe
	# ------------------------------------------------------------------------------------------------
	$nssmExe = $null
	$nssmCandidates = @()

	try
	{
		if ($script:FunctionResults -and $script:FunctionResults.ContainsKey('NssmExe') -and $script:FunctionResults['NssmExe'])
		{
			$nssmCandidates += [string]$script:FunctionResults['NssmExe']
		}
	}
	catch { }

	try
	{
		if ($script:NssmExe) { $nssmCandidates += [string]$script:NssmExe }
	}
	catch { }

	# Preferred default location from your Download_NSSM function
	$nssmCandidates += $script:NssmDefaultPath

	foreach ($p in $nssmCandidates)
	{
		if ([string]::IsNullOrWhiteSpace($p)) { continue }
		if (Test-Path -LiteralPath $p)
		{
			$nssmExe = $p
			break
		}
	}

	# ------------------------------------------------------------------------------------------------
	# Normalize lane map (we only use lane numbers; NEVER build \\POS### paths)
	# ------------------------------------------------------------------------------------------------
	$LaneMapNormalized = $null
	if ($script:FunctionResults.ContainsKey('LaneNumToMachineName') -and $script:FunctionResults['LaneNumToMachineName'])
	{
		$tmp = [ordered]@{ }

		foreach ($kv in $script:FunctionResults['LaneNumToMachineName'].GetEnumerator())
		{
			$k = [string]$kv.Key
			$v = [string]$kv.Value

			$l = $null
			$m = $null

			if ($k -match '^\d{1,3}$')
			{
				$l = ("{0:000}" -f ([int]$k))
				$m = $v
			}
			elseif ($k -match '(?i)(\d{3})$')
			{
				$l = $matches[1]
				$m = if (-not [string]::IsNullOrWhiteSpace($v)) { $v }
				else { $k }
			}

			if (-not $l) { continue }
			if ([string]::IsNullOrWhiteSpace($m)) { continue }
			if ($m -like '@SMSSERVER') { continue }

			if (-not $tmp.Contains($l)) { $tmp[$l] = $m }
		}

		if ($tmp.Count -gt 0) { $LaneMapNormalized = $tmp }
	}

	# ------------------------------------------------------------------------------------------------
	# UI: Run / Install / Uninstall
	# ------------------------------------------------------------------------------------------------
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Remove Archive Bit"
	$form.Size = New-Object System.Drawing.Size(480, 240)
	$form.StartPosition = "CenterScreen"
	$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.Tag = $null

	$label = New-Object System.Windows.Forms.Label
	$label.Text = "Choose an action:`r`n- Run Now (one-time)`r`n- Install Service (always-on, Ter_Load driven)`r`n- Uninstall Service"
	$label.Location = New-Object System.Drawing.Point(20, 15)
	$label.Size = New-Object System.Drawing.Size(440, 60)
	$label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
	$form.Controls.Add($label)

	$btnRunNow = New-Object System.Windows.Forms.Button
	$btnRunNow.Text = "Run Now"
	$btnRunNow.Location = New-Object System.Drawing.Point(10, 95)
	$btnRunNow.Size = New-Object System.Drawing.Size(140, 45)
	$form.Controls.Add($btnRunNow)

	$btnService = New-Object System.Windows.Forms.Button
	$btnService.Text = "Install Service"
	$btnService.Location = New-Object System.Drawing.Point(165, 95)
	$btnService.Size = New-Object System.Drawing.Size(140, 45)
	$form.Controls.Add($btnService)

	$btnUninstall = New-Object System.Windows.Forms.Button
	$btnUninstall.Text = "Uninstall Service"
	$btnUninstall.Location = New-Object System.Drawing.Point(315, 95)
	$btnUninstall.Size = New-Object System.Drawing.Size(140, 45)
	$form.Controls.Add($btnUninstall)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Location = New-Object System.Drawing.Point(190, 160)
	$btnCancel.Size = New-Object System.Drawing.Size(90, 32)
	$form.Controls.Add($btnCancel)

	$btnRunNow.Add_Click({ $form.Tag = "run"; $form.DialogResult = 'OK'; $form.Close() })
	$btnService.Add_Click({ $form.Tag = "service"; $form.DialogResult = 'OK'; $form.Close() })
	$btnUninstall.Add_Click({ $form.Tag = "uninstall"; $form.DialogResult = 'OK'; $form.Close() })
	$btnCancel.Add_Click({ $form.Tag = "cancel"; $form.DialogResult = 'Cancel'; $form.Close() })

	$form.AcceptButton = $btnRunNow
	$form.CancelButton = $btnCancel

	[void]$form.ShowDialog()
	$selectedAction = [string]$form.Tag
	$form.Dispose()

	if ($selectedAction -eq "cancel" -or -not $selectedAction)
	{
		Write_Log "User cancelled Remove_Archive_Bit." "yellow"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# UNINSTALL SERVICE
	# ------------------------------------------------------------------------------------------------
	if ($selectedAction -eq "uninstall")
	{
		$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
			[Security.Principal.WindowsBuiltInRole]::Administrator
		)
		if (-not $isAdmin)
		{
			[void][System.Windows.Forms.MessageBox]::Show(
				"Uninstalling the service requires Administrator rights.`r`n`r`nPlease re-run as Administrator and try again.",
				"Administrator Required",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			Write_Log "ERROR: Uninstall requested without Administrator rights." "red"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		$existingSvc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
		if (-not $existingSvc)
		{
			Write_Log "Service '$serviceName' is not installed." "yellow"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		# Try native uninstall first (no NSSM needed)
		try { & sc.exe stop $serviceName *> $null }
		catch { }
		try { & sc.exe delete $serviceName *> $null }
		catch { }

		Start-Sleep -Milliseconds 700
		if (-not (Get-Service -Name $serviceName -ErrorAction SilentlyContinue))
		{
			Write_Log "Service '$serviceName' removed via sc.exe (no NSSM required)." "green"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		# If sc.exe didn't fully remove it, fall back to NSSM - but ONLY if NSSM exists (NO downloads)
		if (-not $nssmExe -or -not (Test-Path -LiteralPath $nssmExe))
		{
			Write_Log "ERROR: sc.exe did not fully remove the service, and NSSM is missing." "red"
			Write_Log "Expected NSSM at: $script:NssmDefaultPath (or provided by Download_NSSM at startup)" "yellow"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		try { & $nssmExe stop $serviceName *> $null }
		catch { }
		try { & $nssmExe remove $serviceName confirm *> $null }
		catch { }

		Start-Sleep -Milliseconds 700
		if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue)
		{
			Write_Log "ERROR: NSSM attempted removal but service still exists. Check Services.msc for '$serviceName'." "red"
			Write_Log "NSSM used: $nssmExe" "yellow"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		Write_Log "Service '$serviceName' uninstalled successfully." "green"
		Write_Log "NSSM used: $nssmExe" "green"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# INSTALL SERVICE (LOCAL Office + LOCAL Ter_Load, REALTIME WATCHERS)
	# ------------------------------------------------------------------------------------------------
	if ($selectedAction -eq "service")
	{
		$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
			[Security.Principal.WindowsBuiltInRole]::Administrator
		)
		if (-not $isAdmin)
		{
			[void][System.Windows.Forms.MessageBox]::Show(
				"Installing the service requires Administrator rights.`r`n`r`nPlease re-run as Administrator and try again.",
				"Administrator Required",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			Write_Log "ERROR: Service install requested without Administrator rights." "red"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		if (-not $storeNumber)
		{
			Write_Log "ERROR: Store number not present in context." "red"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}
		if (-not (Test-Path -LiteralPath $terFile))
		{
			Write_Log "ERROR: Ter_Load.sql could not be located at $terFile" "red"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		# NSSM must exist for install (NO downloads here)
		if (-not $nssmExe -or -not (Test-Path -LiteralPath $nssmExe))
		{
			Write_Log "ERROR: NSSM is missing. This function will not download it." "red"
			Write_Log "Expected NSSM at: $script:NssmDefaultPath (or provided by Download_NSSM at startup)" "yellow"
			Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
			return
		}

		# Service config
		$configObj = [ordered]@{
			StoreNumber = "$storeNumber"
			OfficePathRaw = "$OfficePath"
			OfficePathLocal = "$OfficePathLocalResolved"
			TerFileLocal = "$terFile"
			LogsDir	    = "$svcLogsDir"
			AutoPollSeconds = [int]$AutoPollSeconds
			LaneMap	    = $LaneMapNormalized
		}
		($configObj | ConvertTo-Json -Depth 7) | Set-Content -LiteralPath $svcConfigPath -Encoding UTF8 -Force
		Write_Log "Service config saved: $svcConfigPath" "green"

		# Service script (LOCAL paths only)
		$svcContent = @"
#requires -Version 5.1
`$ErrorActionPreference = 'Continue'

# --- Load config ---
`$configPath = '$($svcConfigPath.Replace("'", "''"))'
if (-not (Test-Path -LiteralPath `$configPath)) { exit 2 }

try { `$cfg = Get-Content -LiteralPath `$configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
catch { exit 3 }

`$StoreNumber     = [string]`$cfg.StoreNumber
`$OfficePath      = [string]`$cfg.OfficePathLocal
`$OfficePathRaw   = [string]`$cfg.OfficePathRaw
`$TerFile         = [string]`$cfg.TerFileLocal
`$LogsDir         = [string]`$cfg.LogsDir
`$AutoPollSeconds = [int]`$cfg.AutoPollSeconds
if (`$AutoPollSeconds -lt 1) { `$AutoPollSeconds = 3 }

# --- Resolve OfficePath locally again (defensive) ---
try {
	`$op = [string]`$OfficePath
	if ([string]::IsNullOrWhiteSpace(`$op)) { `$op = [string]`$OfficePathRaw }
	if (`$op -match '^(?i)\\\\([^\\]+)\\([^\\]+)(\\.*)?$') {
		`$uncHost = [string]`$matches[1]
		`$share = [string]`$matches[2]
		`$rest = [string]`$matches[3]

		`$isLocalHost = `$false
		if (`$uncHost -match '^(?i)localhost$') { `$isLocalHost = `$true }
		elseif (`$uncHost -eq `$env:COMPUTERNAME) { `$isLocalHost = `$true }
		elseif (`$uncHost -eq ("{0}.{1}" -f `$env:COMPUTERNAME, `$env:USERDNSDOMAIN)) { `$isLocalHost = `$true }

		if (`$isLocalHost) {
			`$rest2 = ""
			if (`$rest) { `$rest2 = `$rest.TrimStart('\') }

			`$sharePath = `$null
			try {
				`$w = Get-CimInstance -ClassName Win32_Share -Filter ("Name='{0}'" -f `$share.Replace("'", "''")) -ErrorAction Stop
				if (`$w -and `$w.Path) { `$sharePath = [string]`$w.Path }
			} catch { }

			if (-not [string]::IsNullOrWhiteSpace(`$sharePath)) {
				if (-not [string]::IsNullOrWhiteSpace(`$rest2)) {
					`$OfficePath = Join-Path `$sharePath `$rest2
				} else {
					`$OfficePath = `$sharePath
				}
			}
		}
	}
} catch { }

if (-not (Test-Path -LiteralPath `$LogsDir)) { New-Item -Path `$LogsDir -ItemType Directory -Force | Out-Null }
`$LogFile = Join-Path `$LogsDir 'Remove_Archive_Bit_Service.log'

try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] Service starting..." -f (Get-Date)) -Encoding UTF8 } catch { }
try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] OfficePath local: {1}" -f (Get-Date), `$OfficePath) -Encoding UTF8 } catch { }
try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] Ter_Load local: {1}" -f (Get-Date), `$TerFile) -Encoding UTF8 } catch { }

`$watchers = New-Object System.Collections.ArrayList
`$terWatcher = `$null
`$needRebuild = `$true

while (`$true)
{
	# Touched item -> clear just its bits. Only a Ter_Load.sql change or a watcher overflow
	# (_e) rebuilds everything (every file event used to trigger a full attrib rescan).
	try {
		`$clearBits = [System.IO.FileAttributes]::Archive -bor [System.IO.FileAttributes]::ReadOnly
		foreach (`$evt in @(Get-Event -ErrorAction SilentlyContinue | Where-Object { [string]`$_.SourceIdentifier -like 'RAB_*' }))
		{
			`$sid = [string]`$evt.SourceIdentifier
			if (`$sid -like 'RAB_TER_WATCH*' -or `$sid -like '*_e') { `$needRebuild = `$true }
			else {
				try {
					`$p = [string]`$evt.SourceEventArgs.FullPath
					if (`$p -and (Test-Path -LiteralPath `$p)) {
						`$it = Get-Item -LiteralPath `$p -Force -ErrorAction Stop
						if (`$it.Attributes -band `$clearBits) {
							`$newAttr = [int](`$it.Attributes) -band (-bnot [int]`$clearBits)
							if (`$newAttr -eq 0) { `$newAttr = [int][System.IO.FileAttributes]::Normal }
							`$it.Attributes = [System.IO.FileAttributes]`$newAttr
						}
					}
				} catch { }
			}
			try { Remove-Event -EventIdentifier `$evt.EventIdentifier -ErrorAction SilentlyContinue } catch { }
		}
	} catch { }

	if (`$needRebuild)
	{
		`$needRebuild = `$false

		# Dispose old watchers + unregister all event ids we create
		try {
			foreach (`$w in @(`$watchers)) {
				try { `$w.EnableRaisingEvents = `$false } catch { }
				try {
					`$id = 'RAB_FSW_' + `$w.GetHashCode().ToString()
					Unregister-Event -SourceIdentifier `$id -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier (`$id + '_c') -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier (`$id + '_r') -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier (`$id + '_d') -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier (`$id + '_e') -ErrorAction SilentlyContinue
				} catch { }
				try { `$w.Dispose() } catch { }
			}
			`$watchers.Clear() | Out-Null
		} catch { }

		# Dispose Ter watcher
		try {
			if (`$terWatcher) {
				try { `$terWatcher.EnableRaisingEvents = `$false } catch { }
				try {
					Unregister-Event -SourceIdentifier 'RAB_TER_WATCH' -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier 'RAB_TER_WATCH_C' -ErrorAction SilentlyContinue
					Unregister-Event -SourceIdentifier 'RAB_TER_WATCH_R' -ErrorAction SilentlyContinue
				} catch { }
				try { `$terWatcher.Dispose() } catch { }
			}
		} catch { }

		# --- Build LOCAL lane roots from Ter_Load.sql (ROBUST: regex extract XF<store><lane>) ---
		`$laneRoots = New-Object System.Collections.Generic.List[string]
		`$lines = @()
		try { `$lines = Get-Content -LiteralPath `$TerFile -ErrorAction SilentlyContinue } catch { }

		`$picked = @()
		try { `$picked = `$lines | Where-Object { `$_ -match 'Terminal\s*0' } } catch { }

		`$text = ""
		if (`$picked -and `$picked.Count -gt 0) { `$text = (`$picked -join "`n") }
		else { `$text = (`$lines -join "`n") }

		if (-not [string]::IsNullOrWhiteSpace(`$text))
		{
			try {
				`$snEsc = [regex]::Escape(`$StoreNumber)
				`$rx = '(?i)XF' + `$snEsc + '\d{3}'

				foreach (`$m in [regex]::Matches(`$text, `$rx))
				{
					`$folder = ([string]`$m.Value).ToUpper()
					if (`$folder.EndsWith('900') -or `$folder.EndsWith('901')) { continue }

					`$local = Join-Path `$OfficePath `$folder
					if (-not `$laneRoots.Contains(`$local)) { [void]`$laneRoots.Add(`$local) }
				}
			} catch { }
		}

		# Fallback: enumerate local XF folders for this store
		if (`$laneRoots.Count -eq 0)
		{
			try {
				foreach (`$d in Get-ChildItem -LiteralPath `$OfficePath -Directory -ErrorAction SilentlyContinue)
				{
					if (`$d.Name -match ('(?i)^XF' + [regex]::Escape(`$StoreNumber) + '\d{3}$'))
					{
						if (`$d.Name -match ('(?i)^XF' + [regex]::Escape(`$StoreNumber) + '(900|901)$')) { continue }
						if (-not `$laneRoots.Contains(`$d.FullName)) { [void]`$laneRoots.Add(`$d.FullName) }
					}
				}
			} catch { }
		}

		# --- Server roots LOCAL ---
		`$serverRoots = New-Object System.Collections.Generic.List[string]
		`$sr900 = Join-Path `$OfficePath ("XF{0}900" -f `$StoreNumber)
		`$sr901 = Join-Path `$OfficePath ("XF{0}901" -f `$StoreNumber)
		[void]`$serverRoots.Add(`$sr900)
		[void]`$serverRoots.Add(`$sr901)

		`$allRoots = @()
		try { `$allRoots += `$laneRoots.ToArray() } catch { }
		try { `$allRoots += `$serverRoots.ToArray() } catch { }

		# --- Full cleanup ---
		try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] Full cleanup starting..." -f (Get-Date)) -Encoding UTF8 } catch { }

		foreach (`$root in `$allRoots)
		{
			if ([string]::IsNullOrWhiteSpace(`$root)) { continue }
			`$r = `$root.Trim().Trim('"')
			if (-not (Test-Path -LiteralPath `$r)) { continue }
			try {
				`$cmd = 'attrib -a -r "{0}\*" /S /D' -f `$r
				& `$env:ComSpec /c `$cmd 2>`$null | Out-Null
			} catch { }
		}

		try {
			Add-Content -LiteralPath `$LogFile -Value ("[{0}] Full cleanup done. LaneRoots={1}, ServerRoots={2}. Building watchers..." -f (Get-Date), `$laneRoots.Count, `$serverRoots.Count) -Encoding UTF8
		} catch { }

		# --- Build filesystem watchers ---
		foreach (`$root in `$allRoots)
		{
			if ([string]::IsNullOrWhiteSpace(`$root)) { continue }
			`$r = `$root.Trim().Trim('"')
			if (-not (Test-Path -LiteralPath `$r)) { continue }

			try {
				`$fsw = New-Object System.IO.FileSystemWatcher
				`$fsw.Path = `$r
				`$fsw.Filter = '*'
				`$fsw.IncludeSubdirectories = `$true
				`$fsw.NotifyFilter = [System.IO.NotifyFilters]'FileName, DirectoryName, LastWrite, Attributes'
				`$fsw.InternalBufferSize = 65536

				`$id = 'RAB_FSW_' + `$fsw.GetHashCode().ToString()
				Register-ObjectEvent -InputObject `$fsw -EventName Created -SourceIdentifier `$id | Out-Null
				Register-ObjectEvent -InputObject `$fsw -EventName Changed -SourceIdentifier (`$id + '_c') | Out-Null
				Register-ObjectEvent -InputObject `$fsw -EventName Renamed -SourceIdentifier (`$id + '_r') | Out-Null
				Register-ObjectEvent -InputObject `$fsw -EventName Deleted -SourceIdentifier (`$id + '_d') | Out-Null
				Register-ObjectEvent -InputObject `$fsw -EventName Error -SourceIdentifier (`$id + '_e') | Out-Null

				`$fsw.EnableRaisingEvents = `$true
				[void]`$watchers.Add(`$fsw)
			} catch { }
		}

		# --- Watch Ter_Load.sql itself ---
		try {
			`$terDir = Split-Path -Parent `$TerFile
			if (Test-Path -LiteralPath `$terDir)
			{
				`$terWatcher = New-Object System.IO.FileSystemWatcher
				`$terWatcher.Path = `$terDir
				`$terWatcher.Filter = (Split-Path -Leaf `$TerFile)
				`$terWatcher.IncludeSubdirectories = `$false
				`$terWatcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite, Size'

				Register-ObjectEvent -InputObject `$terWatcher -EventName Changed -SourceIdentifier 'RAB_TER_WATCH' | Out-Null
				Register-ObjectEvent -InputObject `$terWatcher -EventName Created -SourceIdentifier 'RAB_TER_WATCH_C' | Out-Null
				Register-ObjectEvent -InputObject `$terWatcher -EventName Renamed -SourceIdentifier 'RAB_TER_WATCH_R' | Out-Null

				`$terWatcher.EnableRaisingEvents = `$true
			}
		} catch { }

		try { Add-Content -LiteralPath `$LogFile -Value ("[{0}] Watchers active: {1}" -f (Get-Date), `$watchers.Count) -Encoding UTF8 } catch { }
	}

	# Wake on the next event (or after the poll interval)
	`$null = Wait-Event -Timeout `$AutoPollSeconds
}
"@

		$svcContent | Set-Content -LiteralPath $svcPS1 -Encoding UTF8 -Force
		Write_Log "Service script saved: $svcPS1" "green"

		$serviceStdOut = Join-Path $svcLogsDir "Remove_Archive_Bit_Service_stdout.log"
		$serviceStdErr = Join-Path $svcLogsDir "Remove_Archive_Bit_Service_stderr.log"

		$existingSvc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
		if ($existingSvc)
		{
			try { & $nssmExe stop $serviceName *> $null }
			catch { }
			try { & $nssmExe remove $serviceName confirm *> $null }
			catch { }
		}

		& $nssmExe install $serviceName "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
				   -NoProfile -ExecutionPolicy Bypass -File "$svcPS1" *> $null

		& $nssmExe set $serviceName AppDirectory "$svcScripts" *> $null
		& $nssmExe set $serviceName AppStdout "$serviceStdOut" *> $null
		& $nssmExe set $serviceName AppStderr "$serviceStdErr" *> $null
		& $nssmExe set $serviceName Start SERVICE_AUTO_START *> $null

		# LOCAL filesystem permissions; avoids UNC auth failures
		& $nssmExe set $serviceName ObjectName "LocalSystem" "" *> $null

		& $nssmExe set $serviceName DisplayName "Remove Archive Bit (LOCAL XF Watcher)" *> $null
		& $nssmExe set $serviceName Description "Clears archive/read-only bits in LOCAL XF folders; watches Ter_Load.sql and filesystem changes." *> $null

		& $nssmExe start $serviceName *> $null

		Write_Log "Service '$serviceName' installed and started as: LocalSystem" "green"
		Write_Log "Logs: $svcLogsDir" "green"
		Write_Log "NSSM used: $nssmExe" "green"

		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# ------------------------------------------------------------------------------------------------
	# RUN NOW (one-time)
	# ------------------------------------------------------------------------------------------------
	if (-not $storeNumber)
	{
		Write_Log "ERROR: Store number not present in context." "red"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}
	if (-not (Test-Path -LiteralPath $terFile))
	{
		Write_Log "ERROR: Ter_Load.sql could not be located at $terFile" "red"
		Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
		return
	}

	# --- Run Now: ALWAYS LOCAL paths (never \\POS###)
	$didLaneWork = $false

	if ($LaneMapNormalized)
	{
		foreach ($laneNum in ($LaneMapNormalized.Keys | Sort-Object { [int]$_ }))
		{
			$localLanePath = Join-Path $OfficePathLocalResolved ("XF{0}{1}" -f $storeNumber, $laneNum)
			$lp = $localLanePath.Trim().Trim('"')

			if (Test-Path -LiteralPath $lp)
			{
				Write_Log "Refreshing attributes (recursive) in $lp..." "green"
				$cmd = 'attrib -a -r "{0}\*" /S /D' -f $lp
				try { & $env:ComSpec /c $cmd 2>$null | Out-Null }
				catch { Write_Log "ERROR: Failed attrib on: $lp" "red" }
				$didLaneWork = $true
			}
			else
			{
				Write_Log "Local lane path not found (skip): $lp" "yellow"
			}
		}
	}
	else
	{
		Write_Log "Lane map not found/usable. Falling back to Ter_Load.sql parsing (LOCAL rebuild)..." "yellow"
	}

	# Fallback: robust Ter_Load scan for XF<store><lane>
	if (-not $didLaneWork)
	{
		$lanePaths = New-Object System.Collections.Generic.List[string]
		$lines = @()
		try { $lines = Get-Content -LiteralPath $terFile -ErrorAction SilentlyContinue }
		catch { }

		$picked = @()
		try { $picked = $lines | Where-Object { $_ -match 'Terminal\s*0' } }
		catch { }

		$text = ""
		if ($picked -and $picked.Count -gt 0) { $text = ($picked -join "`n") }
		else { $text = ($lines -join "`n") }

		if (-not [string]::IsNullOrWhiteSpace($text))
		{
			try
			{
				$snEsc = [regex]::Escape($storeNumber)
				$rx = '(?i)XF' + $snEsc + '\d{3}'

				foreach ($m in [regex]::Matches($text, $rx))
				{
					$folder = ([string]$m.Value).ToUpper()
					if ($folder.EndsWith('900') -or $folder.EndsWith('901')) { continue }

					$local = Join-Path $OfficePathLocalResolved $folder
					if (-not $lanePaths.Contains($local)) { [void]$lanePaths.Add($local) }
				}
			}
			catch { }
		}

		# fallback: enumerate folders if still empty
		if ($lanePaths.Count -eq 0)
		{
			try
			{
				foreach ($d in Get-ChildItem -LiteralPath $OfficePathLocalResolved -Directory -ErrorAction SilentlyContinue)
				{
					if ($d.Name -match ('(?i)^XF' + [regex]::Escape($storeNumber) + '\d{3}$'))
					{
						if ($d.Name -match ('(?i)^XF' + [regex]::Escape($storeNumber) + '(900|901)$')) { continue }
						if (-not $lanePaths.Contains($d.FullName)) { [void]$lanePaths.Add($d.FullName) }
					}
				}
			}
			catch { }
		}

		if ($lanePaths.Count -gt 0)
		{
			foreach ($p in $lanePaths)
			{
				$pp = $p.Trim().Trim('"')
				if (Test-Path -LiteralPath $pp)
				{
					Write_Log "Refreshing attributes (recursive) in $pp..." "green"
					$cmd = 'attrib -a -r "{0}\*" /S /D' -f $pp
					try { & $env:ComSpec /c $cmd 2>$null | Out-Null }
					catch { Write_Log "ERROR: Failed attrib on: $pp" "red" }
				}
				else
				{
					Write_Log "Path not found (skip): $pp" "yellow"
				}
			}
		}
		else
		{
			Write_Log "No LOCAL lane paths could be derived from Ter_Load.sql and no XF lane folders found locally." "yellow"
		}
	}

	# Server roots LOCAL
	foreach ($suffix in 900, 901)
	{
		$serverPath = Join-Path $OfficePathLocalResolved ("XF{0}{1}" -f $storeNumber, $suffix)
		$sp = $serverPath.Trim().Trim('"')
		if (Test-Path -LiteralPath $sp)
		{
			Write_Log "Refreshing attributes (recursive) in $sp" "green"
			$cmd = 'attrib -a -r "{0}\*" /S /D' -f $sp
			try { & $env:ComSpec /c $cmd 2>$null | Out-Null }
			catch { Write_Log "ERROR: Failed attrib on: $sp" "red" }
		}
		else
		{
			Write_Log "Server path not found: $sp" "yellow"
		}
	}

	Write_Log "`r`n==================== Remove_Archive_Bit Function Completed ====================" "blue"
}

# ===================================================================================================
#          SECTION: More ported core (remote lanes, scales, SMS functions, exports - verbatim)
# ---------------------------------------------------------------------------------------------------
#   Source: Desktop\script.ps1 (Alex_C.T). Business logic UNCHANGED. Background-job helpers, XE folder
#   cleaner, lane time sync, SMS key, Install_FUNCTIONS_Into_SMS, scale test item / ScaleCommApp
#   troubleshooter, VNC export + passwords, remote machine info.
# ===================================================================================================

# ===================================================================================================
#                              FUNCTION: Get_Background_Function_Job_Result
# ---------------------------------------------------------------------------------------------------
# Description:
#   Polls a job created by Start_Background_Function_Job. Returns $null while the job is still
#   running. When it has finished, receives the result (before removing the job), merges the
#   job's FunctionResults and ExtraVars back into the live script scope and returns the
#   result object. Safe to call repeatedly.
#
# Parameters:
#   - Job  [Job]: the job object returned by Start_Background_Function_Job
#
# Variables:
#   - $script:BgfJobs: the job's registry entry is removed on completion
#
# Returns:
#   The job result object (Success/FunctionResults/ExtraVars/Error), or $null.
# ===================================================================================================

function Get_Background_Function_Job_Result
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.Job]$Job
	)

	if (-not $Job) { return $null }
	if ($Job.State -in @('Running', 'NotStarted')) { return $null }

	$result = $null
	try { $result = @(Receive-Job -Job $Job -ErrorAction SilentlyContinue) | Select-Object -Last 1 } catch { }
	try { Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue | Out-Null } catch { }

	if ($result)
	{
		if ($result.FunctionResults)
		{
			if (-not $script:FunctionResults) { $script:FunctionResults = @{} }
			foreach ($k in @($result.FunctionResults.Keys))
			{
				$script:FunctionResults[$k] = $result.FunctionResults[$k]
			}
		}
		if ($result.ExtraVars)
		{
			foreach ($k in @($result.ExtraVars.Keys))
			{
				Set-Variable -Scope Script -Name $k -Value $result.ExtraVars[$k] -Force
			}
		}
	}

	if ($script:BgfJobs -and $script:BgfJobs.ContainsKey($Job.Id))
	{
		$script:BgfJobs.Remove($Job.Id)
	}
	return $result
}

# ===================================================================================================
#                              FUNCTION: Start_Background_Function_Job
# ---------------------------------------------------------------------------------------------------
# Description:
#   The one background-job builder for this script (v2.6.9). Runs a script function in a
#   background job without blocking the UI thread. The job receives the function code,
#   Write_Log, any dependent functions, a copy of the FunctionResults bag and seed variables,
#   and its results are merged back into the live script scope when it finishes.
#
# Parameters:
#   - FunctionName     [string]: function to run in the job (required)
#   - Label            [string]: shown in the status bar / logged while the job runs
#   - CallArgs         [hashtable]: named arguments splatted onto the function call
#   - SeedVars         [hashtable]: script variables the function reads (name -> value)
#   - ReturnVarNames   [string[]]: script variables the function writes back (names only)
#   - DependentBodies  [hashtable]: extra functions the target calls (name -> body text)
#   - OnComplete       [scriptblock]: run on the UI thread when the job completes
#   - NotifyOnComplete [switch]: box line when the job ends without a result
#
# Variables:
#   - $script:BgfJobs: registry of running jobs for the shared reaper timer
#
# Returns:
#   The Start-Job job object, or $null if the function was not found.
# ===================================================================================================

function Start_Background_Function_Job
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$FunctionName,
		# shown in the status bar / logged while the job runs
		[string]$Label = '',
		# named arguments splatted onto the function call, e.g. @{ RefreshCache = $true }
		[hashtable]$CallArgs = @{},
		# extra script variables the function reads (name -> value), e.g. @{ SetupFiles = $script:SetupFiles }
		[hashtable]$SeedVars = @{},
		# script variables the function WRITES that must be handed back to the
		# main session (names only), e.g. @('LaneSQLScript', 'ServerSQLScript')
		[string[]]$ReturnVarNames = @(),
		# extra functions the target calls internally (name -> body text).
		# Write_Log is always passed, so the job can log like the main session.
		[hashtable]$DependentBodies = @{},
		# scriptblock run on the UI thread when the job completes; it receives
		# the result object (possibly $null). Used by warmup steps that must
		# hydrate state from the warm cache after the job's network work.
		[scriptblock]$OnComplete = $null,
		# mark the job as one whose abnormal finish ("ended without a result")
		# is worth a line in the log BOX. Normal success lines always go to the
		# file log only - the status bar covers progress, the box sees failures.
		[switch]$NotifyOnComplete
	)

	if (-not $CallArgs) { $CallArgs = @{} }
	if (-not $SeedVars) { $SeedVars = @{} }
	if (-not $ReturnVarNames) { $ReturnVarNames = @() }
	if (-not $DependentBodies) { $DependentBodies = @{} }
	if (-not $script:BgfJobs) { $script:BgfJobs = @{} }

	# NOTE: use Get-Command, NOT ${function:$FunctionName} - the variable-drive
	# syntax with a dynamic name does not resolve from inside a function body
	# under Windows PowerShell 5.1 (only literal names work there).
	$fnCmd = $null
	try { $fnCmd = Get-Command -Name $FunctionName -CommandType Function -ErrorAction SilentlyContinue } catch { }
	if (-not $fnCmd)
	{
		# NOTE: Out-Null keeps the return value clean even if a Write_Log
		# variant ever writes to the pipeline.
		try { Write_Log ("Start_Background_Function_Job: function '{0}' not found, cannot start job" -f $FunctionName) 'red' | Out-Null } catch { }
		return $null
	}
	$fnBody = $fnCmd.ScriptBlock.ToString()

	# Copy the results bag - the job must never mutate the live one. UI objects (the LogBox
	# entry, timers) are left out: they cannot be used in another runspace, serialising a
	# WinForms control into the job is slow, and the deserialised copy would be merged back
	# over the live entry when the job completes.
	$seedResults = @{}
	if ($script:FunctionResults)
	{
		foreach ($k in @($script:FunctionResults.Keys))
		{
			$v = $script:FunctionResults[$k]
			if ($v -is [System.ComponentModel.Component]) { continue }
			$seedResults[$k] = $v
		}
	}

	# Standard context every job needs so its Write_Log can find the file-log
	# folder and honour test mode. Callers do not have to remember this.
	$autoSeed = @{
		ScriptsLog        = $script:ScriptsLog
		IsTestEnvironment = $script:IsTestEnvironment
	}

	$job = Start-Job -Name ('bgf_' + $FunctionName) -ScriptBlock {
		param (
			[string]$FnName,
			[string]$FnBody,
			[hashtable]$Deps,
			[hashtable]$SeedResults,
			[hashtable]$SeedVars,
			[hashtable]$CallArgs,
			[string[]]$ReturnVarNames,
			[string]$WriteLogBody,
			[hashtable]$AutoSeed
		)

		Add-Type -AssemblyName System.Windows.Forms 2>$null
		Add-Type -AssemblyName System.Drawing 2>$null

		Set-Item -Path function:Write_Log -Value ([scriptblock]::Create($WriteLogBody))
		foreach ($depName in @($Deps.Keys))
		{
			Set-Item -Path ('function:' + $depName) -Value ([scriptblock]::Create($Deps[$depName]))
		}
		Set-Item -Path ('function:' + $FnName) -Value ([scriptblock]::Create($FnBody))

		$script:FunctionResults = @{ }
		foreach ($k in @($SeedResults.Keys)) { $script:FunctionResults[$k] = $SeedResults[$k] }
		foreach ($k in @($SeedVars.Keys))
		{
			Set-Variable -Scope Script -Name $k -Value $SeedVars[$k]
		}
		foreach ($k in @($AutoSeed.Keys))
		{
			if (-not (Get-Variable -Scope Script -Name $k -ValueOnly -ErrorAction SilentlyContinue))
			{
				Set-Variable -Scope Script -Name $k -Value $AutoSeed[$k]
			}
		}

		try
		{
			& (Get-Item -Path ('function:' + $FnName)) @CallArgs
			$extraVars = @{}
			foreach ($vn in $ReturnVarNames)
			{
				$vv = Get-Variable -Scope Script -Name $vn -ValueOnly -ErrorAction SilentlyContinue
				if ($null -ne $vv) { $extraVars[$vn] = $vv }
			}
			[pscustomobject]@{
				Success         = $true
				FunctionResults = $script:FunctionResults
				ExtraVars       = $extraVars
				Error           = $null
			}
		}
		catch
		{
			[pscustomobject]@{
				Success         = $false
				FunctionResults = $script:FunctionResults
				ExtraVars       = @{}
				Error           = $_.Exception.Message
			}
		}
	} -ArgumentList @(
		$FunctionName,
		$fnBody,
		$DependentBodies,
		$seedResults,
		$SeedVars,
		$CallArgs,
		$ReturnVarNames,
		${function:Write_Log}.ToString(),
		$autoSeed
	)

	if (-not $job)
	{
		Write_Log ("Start_Background_Function_Job: failed to start job for '{0}'" -f $FunctionName) 'red'
		return $null
	}

	$script:BgfJobs[$job.Id] = @{
		Job              = $job
		Label            = $Label
		Started          = Get-Date
		OnComplete       = $OnComplete
		NotifyOnComplete = [bool]$NotifyOnComplete
	}

	# v2.6.9: ONE shared reaper polls every background FUNCTION job on the
	# UI thread (1s tick), merges finished results and fires the job's
	# OnComplete callback (warmup hydration) or a one-line completion notice.
	# In hosts without a message pump (e.g. the selftest's child process) it
	# simply never ticks, and callers may poll Get_Background_Function_Job_Result
	# directly instead.
	# NOTE: the other timers in this script (ScaleCredReaper, VncExportReaper,
	# RemoteMachineInfoReaper, StartupScaleCredTimer, svcTimer) each manage an
	# INLINE job pool with local-variable closures and specialized collect
	# logic; they deliberately do not go through this engine. Migrating them
	# would mean restructuring those functions into the worker pattern -
	# considered and deferred (no performance gain, real regression risk).
	# The reaper stops itself when no job is left; every new job (re)starts it.
	try
	{
		if (-not $script:BgfJobReaper)
		{
			$script:BgfJobReaper = New-Object System.Windows.Forms.Timer
			$script:BgfJobReaper.Interval = 1000
			$script:BgfJobReaper.Add_Tick({
					try
					{
						if (-not $script:BgfJobs -or $script:BgfJobs.Count -eq 0) { $script:BgfJobReaper.Stop(); return }
						foreach ($key in @($script:BgfJobs.Keys))
						{
							$entry = $script:BgfJobs[$key]
							if (-not $entry -or -not $entry.Job) { continue }
							if ($entry.Job.State -in @('Running', 'NotStarted')) { continue }

							$result = $null
							try { $result = Get_Background_Function_Job_Result -Job $entry.Job } catch { }

							$secs = 0
							try { $secs = ((Get-Date) - $entry.Started).TotalSeconds } catch { }
							if ($result)
							{
								if ($result.Success)
								{
									# v2.6.9: success notices go to the FILE LOG ONLY (-FileOnly).
									# The status bar already shows what the app is doing in the
									# background, so the box stays quiet on success and only
									# sees failures (user preference).
									Write_Log ("Background job '{0}' finished in {1:N1}s" -f $entry.Label, $secs) 'blue' -FileOnly
								}
								else
								{
									Write_Log ("Background job '{0}' failed: {1}" -f $entry.Label, $result.Error) 'red'
								}
							}
							elseif ($entry.NotifyOnComplete)
							{
								Write_Log ("Background job '{0}' ended without a result" -f $entry.Label) 'yellow'
							}

							if ($entry.OnComplete)
							{
								try { & $entry.OnComplete $result } catch { }
							}
						}
					}
					catch { }
			})
		}
		if (-not $script:BgfJobReaper.Enabled) { $script:BgfJobReaper.Start() }
	}
	catch { }
	return $job
}

# ===================================================================================================
#                              FUNCTION: Test-Hosts_Parallel
# ---------------------------------------------------------------------------------------------------
# Description:
#   Probes a list of hostnames/IPs in parallel with a per-probe timeout and returns
#   target -> 'OK' / 'FAIL' / 'ERROR'. Names are resolved with async DNS, then every address
#   is pinged with Ping.SendPingAsync, all in flight together (in batches), so a store full of
#   dead hosts costs about one timeout instead of one timeout per host.
#
# Parameters:
#   - Targets        [string[]]: hostnames or IPv4 addresses to probe (required, pipeline)
#   - TimeoutMs      [int]: per-probe timeout in milliseconds (default 400)
#   - MaxConcurrency [int]: max probes in flight per batch (default 64)
#
# Returns:
#   Hashtable target -> 'OK' (reachable), 'FAIL' (no reply in time) or 'ERROR' (name did
#   not resolve / probe failed).
#
# Notes:
#   - The previous Parallel.ForEach version ran a PowerShell scriptblock on thread-pool
#     threads, which have no runspace: it always threw and fell back to a serial
#     Test-Connection loop (8 dead hosts = 32 seconds).
# ===================================================================================================

function Test-Hosts_Parallel
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[string[]]$Targets,
		[int]$TimeoutMs = 400,
		[int]$MaxConcurrency = 64
	)

	begin { $all = New-Object System.Collections.Generic.List[string] }
	process { foreach ($t in $Targets) { if (-not [string]::IsNullOrWhiteSpace($t) -and -not $all.Contains($t)) { $all.Add($t) } } }
	end
	{
		$out = @{ }
		if ($all.Count -eq 0) { return $out }

		# Wait for a set of tasks without throwing (WaitAll throws when any task faulted)
		$waitTasks = {
			param ($Tasks, [int]$Ms)
			if (@($Tasks).Count -eq 0) { return }
			try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]@($Tasks), $Ms) }
			catch { }
		}

		$batchSize = [Math]::Max(1, $MaxConcurrency)
		for ($start = 0; $start -lt $all.Count; $start += $batchSize)
		{
			$batch = $all.GetRange($start, [Math]::Min($batchSize, $all.Count - $start))

			# 1) Resolve names (IPs are used as-is)
			$address = @{ }
			$dns = @{ }
			foreach ($t in $batch)
			{
				$ip = $null
				if ([System.Net.IPAddress]::TryParse($t, [ref]$ip)) { $address[$t] = $ip }
				else
				{
					try { $dns[$t] = [System.Net.Dns]::GetHostAddressesAsync($t) }
					catch { $out[$t] = 'ERROR' }
				}
			}
			& $waitTasks $dns.Values ([Math]::Max(1500, $TimeoutMs * 3))
			foreach ($t in $dns.Keys)
			{
				$task = $dns[$t]
				$ipv4 = $null
				if ($task.Status -eq 'RanToCompletion') { $ipv4 = $task.Result | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1 }
				if ($ipv4) { $address[$t] = $ipv4 }
				else { $out[$t] = 'ERROR' }
			}

			# 2) Ping every resolved address at once (one Ping object per probe)
			$pings = @{ }
			foreach ($t in $address.Keys)
			{
				$p = New-Object System.Net.NetworkInformation.Ping
				try { $pings[$t] = @{ Ping = $p; Task = $p.SendPingAsync($address[$t], $TimeoutMs) } }
				catch { $p.Dispose(); $out[$t] = 'ERROR' }
			}
			& $waitTasks @($pings.Values | ForEach-Object { $_.Task }) ($TimeoutMs + 1000)
			foreach ($t in $pings.Keys)
			{
				$task = $pings[$t].Task
				if ($task.Status -eq 'RanToCompletion') { $out[$t] = if ($task.Result.Status -eq 'Success') { 'OK' } else { 'FAIL' } }
				elseif ($task.Status -eq 'Faulted') { $out[$t] = 'ERROR' }
				else { $out[$t] = 'FAIL' }
				try { $pings[$t].Ping.Dispose() }
				catch { }
			}
		}

		return $out
	}
}

# ===================================================================================================
#                              FUNCTION: Clear_XE_Folder
# ---------------------------------------------------------------------------------------------------
# Description:
#   Keeps the lane XE (urgent messages) folder clean. At startup it deletes everything except
#   valid S*.??? health files (also clearing old FATAL* errors), then starts a background
#   watcher job that keeps newly arriving FATAL* files and deletes everything else except
#   current health files on a regular sweep.
#
# Parameters:
#   - FolderPath           [string]: XE folder (default <OfficePath>\XE<Store>901)
#   - SweepIntervalSeconds [int]: periodic sweep interval in seconds, 2-3600 (default 30)
#   - HealthMaxAgeDays     [int]: max age in days of health files kept, 1-3650 (default 30)
#
# Returns:
#   The watcher job when started; $null if the folder is missing or the job fails.
#
# Writes:
#   - Deletes files under the XE folder (everything except valid health files)
# ===================================================================================================

function Clear_XE_Folder
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$FolderPath = "$OfficePath\XE${StoreNumber}901",
		[Parameter(Mandatory = $false)]
		[ValidateRange(2, 3600)]
		[int]$SweepIntervalSeconds = 30,
		[Parameter(Mandatory = $false)]
		[ValidateRange(1, 3650)]
		[int]$HealthMaxAgeDays = 30
	)

	# ------------------------------------------------------------------------------------------------
	# 0) Validate folder exists
	# ------------------------------------------------------------------------------------------------
	if (-not (Test-Path -LiteralPath $FolderPath))
	{
		Write_Log "Folder '$FolderPath' (Urgent Messages) does not exist." "red"
		return $null
	}

	# ------------------------------------------------------------------------------------------------
	# 1) STARTUP CLEANUP
	#    - Deletes EVERYTHING except valid S*.??? health files (FATAL* old errors too)
	# ------------------------------------------------------------------------------------------------
	# A file worth keeping: a current lane health report ("This application is not running")
	# for this store, S*.<lane> named after the lane in its From: line, newer than HealthMaxAgeDays.
	# Shared by the startup cleanup and the watcher job (passed to it as text).
	$isHealthFile = {
		param ($Item, [string]$StoreNumber, [int]$HealthMaxAgeDays)
		if ($Item.PSIsContainer -or $Item.Name -notmatch '^S.*\.\w{3}$') { return $false }
		if (((Get-Date) - $Item.LastWriteTime).TotalDays -gt $HealthMaxAgeDays) { return $false }
		$content = $null
		try { $content = Get-Content -LiteralPath $Item.FullName -ErrorAction Stop }
		catch { return $false }
		if (-not $content) { return $false }
		$fromLine = $content | Where-Object { $_ -like 'From:*' } | Select-Object -First 1
		$subjectLine = $content | Where-Object { $_ -like 'Subject:*' } | Select-Object -First 1
		$msgLine = $content | Where-Object { $_ -like 'MSG:*' } | Select-Object -First 1
		$lastStatusLine = $content | Where-Object { $_ -like 'Last recorded status:*' } | Select-Object -First 1
		if ($fromLine -notmatch 'From:\s*(\d{3})(\d{3})') { return $false }
		$fileStoreNumber = $Matches[1]
		$fileLaneNumber = $Matches[2]
		return ($fileStoreNumber -eq $StoreNumber -and
			$Item.Name -match '^S.*\.(\d{3})$' -and $fileLaneNumber -eq $Matches[1] -and
			$subjectLine -match 'Subject:\s*Health' -and
			$msgLine -match 'MSG:\s*This application is not running\.' -and
			$lastStatusLine -match 'Last recorded status:\s*[\d\s:,-]+TRANS,(\d+)')
	}

	try
	{
		$items = Get-ChildItem -LiteralPath $FolderPath -Recurse -Force -ErrorAction SilentlyContinue |
		Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 }

		foreach ($item in $items)
		{
			# Startup mode: FATAL* (old errors) and folders go too
			$keep = ($item.Name -notlike 'FATAL*') -and (& $isHealthFile $item $StoreNumber $HealthMaxAgeDays)
			if ($keep) { continue }
			for ($i = 0; $i -lt 8; $i++)
			{
				try
				{
					if (Test-Path -LiteralPath $item.FullName) { Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction Stop }
					break
				}
				catch { Start-Sleep -Milliseconds 150 }
			}
		}
	}
	catch
	{
		Write_Log "An error occurred during startup cleaning of '$FolderPath': $_" "red"
	}
	# ------------------------------------------------------------------------------------------------
	# 2) MONITOR (FileSystemWatcher in a background job)
	#    - Keeps NEW FATAL* while running
	#    - Deletes everything else except valid health files
	# ------------------------------------------------------------------------------------------------
	try
	{
		$jobName = "XEWatcherJob_$StoreNumber"

		# Replace any existing watcher job
		$existing = Get-Job | Where-Object { $_.Name -eq $jobName } # -Name on a missing job records an error
		if ($existing)
		{
			try { Stop-Job -Job $existing -Force -ErrorAction SilentlyContinue | Out-Null }
			catch { }
			try { Remove-Job -Job $existing -Force -ErrorAction SilentlyContinue | Out-Null }
			catch { }
		}

		$job = Start-Job -Name $jobName -ArgumentList $FolderPath, $StoreNumber, $HealthMaxAgeDays, $SweepIntervalSeconds, $isHealthFile.ToString() -ScriptBlock {
			param (
				[string]$FolderPath,
				[string]$StoreNumber,
				[int]$HealthMaxAgeDays,
				[int]$SweepIntervalSeconds,
				[string]$IsHealthFileText
			)
			$IsHealthFile = [scriptblock]::Create($IsHealthFileText)

			$ErrorActionPreference = 'SilentlyContinue'

			# --- Watcher object ---
			$fsw = New-Object System.IO.FileSystemWatcher
			$fsw.Path = $FolderPath
			$fsw.IncludeSubdirectories = $true
			$fsw.NotifyFilter = [IO.NotifyFilters]'FileName, DirectoryName, LastWrite, CreationTime'

			# --- Shared processor (MONITOR MODE: keep FATAL*) ---
			$KeptCache = @{ }   # path -> @{ Stamp; Checked } of health files judged "keep"
			$ProcessPath = {
				param ([string]$Path)

				if (-not $Path) { return }

				# Retry Get-Item because events can fire before writer finishes.
				$item = $null
				for ($t = 0; $t -lt 8; $t++)
				{
					try
					{
						if (-not (Test-Path -LiteralPath $Path)) { return }
						$item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
						break
					}
					catch { Start-Sleep -Milliseconds 120 }
				}
				if (-not $item) { return }

				# Keep NEW FATAL* while running
				if ($item.Name -like 'FATAL*') { return }

				# Skip reparse points
				if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return }

				# A health file already judged "keep" and unchanged since is not re-read on every
				# sweep / event (re-checked hourly so the age limit still applies)
				$stamp = $item.LastWriteTimeUtc.Ticks
				$prev = $KeptCache[$item.FullName]
				if ($prev -and $prev.Stamp -eq $stamp -and ((Get-Date) - $prev.Checked).TotalHours -lt 1) { return }

				$keep = & $IsHealthFile $item $StoreNumber $HealthMaxAgeDays

				if ($keep) { $KeptCache[$item.FullName] = @{ Stamp = $stamp; Checked = Get-Date } }
				else
				{
					[void]$KeptCache.Remove($item.FullName)
					for ($i = 0; $i -lt 8; $i++)
					{
						try
						{
							if (Test-Path -LiteralPath $item.FullName)
							{
								Remove-Item -LiteralPath $item.FullName -Force -Recurse -ErrorAction Stop
							}
							break
						}
						catch { Start-Sleep -Milliseconds 150 }
					}
				}
			}

			# Register events (no -Action) and consume them with Wait-Event (best in jobs)
			$createdId = "XE_Created_$StoreNumber"
			$changedId = "XE_Changed_$StoreNumber"
			$renamedId = "XE_Renamed_$StoreNumber"

			Register-ObjectEvent -InputObject $fsw -EventName Created -SourceIdentifier $createdId | Out-Null
			Register-ObjectEvent -InputObject $fsw -EventName Changed -SourceIdentifier $changedId | Out-Null
			Register-ObjectEvent -InputObject $fsw -EventName Renamed -SourceIdentifier $renamedId | Out-Null

			$fsw.EnableRaisingEvents = $true
			$nextSweep = (Get-Date).AddSeconds($SweepIntervalSeconds)

			while ($true)
			{
				# Wait briefly for an event
				$e = Wait-Event -Timeout 1

				if ($e)
				{
					try
					{
						$path = $e.SourceEventArgs.FullPath
						& $ProcessPath -Path $path
					}
					catch { }
					finally
					{
						# Always remove handled event from queue
						try { Remove-Event -EventIdentifier $e.EventIdentifier -ErrorAction SilentlyContinue }
						catch { }
					}

					# Drain burst
					while ($true)
					{
						$e2 = Wait-Event -Timeout 0
						if (-not $e2) { break }
						try
						{
							$path2 = $e2.SourceEventArgs.FullPath
							& $ProcessPath -Path $path2
						}
						catch { }
						finally
						{
							try { Remove-Event -EventIdentifier $e2.EventIdentifier -ErrorAction SilentlyContinue }
							catch { }
						}
					}
				}

				# Periodic sweep fallback
				if ((Get-Date) -ge $nextSweep)
				{
					$nextSweep = (Get-Date).AddSeconds($SweepIntervalSeconds)

					try
					{
						if (Test-Path -LiteralPath $FolderPath)
						{
							Get-ChildItem -LiteralPath $FolderPath -Recurse -Force -ErrorAction SilentlyContinue |
							Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 } |
							ForEach-Object { & $ProcessPath -Path $_.FullName }
						}
					}
					catch { }
				}
			}
		}

		return $job
	}
	catch
	{
		Write_Log "Failed to start XE watcher job: $_" "red"
		return $null
	}
}

# ===================================================================================================
#                              FUNCTION: Send_TIME_Sync_Lanes_Worker
# ---------------------------------------------------------------------------------------------------
# Description:
#   The per-lane time-sync work behind Send_SERVER_time_to_Lanes (v2.6.9): builds a fresh
#   WIZRPL date/time string at start and pushes a time sync (remote schtasks net time task)
#   to each selected lane. Split out of the wizard so it can run in a background job via
#   Start_Background_Function_Job; running it inline is still supported as a fallback.
#
# Parameters:
#   - StoreNumber     [string]: store to process (required)
#   - SelectedLanes   [any]: lanes to sync, numbers or lane objects (required)
#   - IsScheduling    [switch]: create a recurring task instead of a one-shot sync
#   - IntervalMinutes [int]: schedule interval when -IsScheduling (default 0)
#
# Variables:
#   - $script:FunctionResults['LaneNumToMachineName']: lane map (input)
#
# Writes:
#   - remote scheduled time-sync tasks on the selected lanes
# ===================================================================================================

function Send_TIME_Sync_Lanes_Worker
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $true)]
		$SelectedLanes,
		[switch]$IsScheduling,
		[int]$IntervalMinutes = 0
	)

	# The mailslot type does not exist in a job runspace; this is a
	# no-op where it is already loaded.
	try { Ensure_MailslotSender_Type } catch { }

	# 3. Build Time String
	$now = Get-Date
	$dateStr = $now.ToString("MM/dd/yyyy")
	$timeStr = $now.ToString("HHmmss")
	$wizCommand = "@WIZRPL(DATE=$dateStr)@WIZRPL(TIME=$timeStr)"

	# This server's IP for 'net time' (resolved once, not per lane)
	$serverIP = $env:COMPUTERNAME
	try { $serverIP = (Test-Connection -ComputerName $env:COMPUTERNAME -Count 1 -ErrorAction Stop).IPv4Address.IPAddressToString }
	catch { }
	$netTimeCommand = "net time \\$serverIP /set /yes"

	foreach ($lane in $SelectedLanes)
	{
		# Lane number normalization
		$laneNum = if ($lane -is [pscustomobject] -and $lane.LaneNumber) { $lane.LaneNumber }
		elseif ($lane -is [int]) { "{0:D3}" -f $lane }
		elseif ($lane.Length -lt 3) { $lane.PadLeft(3, '0') }
		else { $lane }

		$machine = $script:FunctionResults['LaneNumToMachineName'][$laneNum]
		if (-not $machine)
		{
			Write_Log "No machine name found for lane $laneNum." "red"
			continue
		}

		$success = $false

		# -------- PRIMARY: REMOTE SCHTASKS (SYSTEM task running net time against this server) --------
		try
		{
			$taskName = if ($IsScheduling) { "ScheduledTimeSync" } else { "SyncTime" }
			$createArgs = @('/create', '/s', $machine, '/tn', $taskName, '/tr', $netTimeCommand, '/ru', 'SYSTEM', '/f', '/rl', 'HIGHEST')
			if ($IsScheduling) { $createArgs += @('/sc', 'minute', '/mo', [string]$IntervalMinutes) }
			else
			{
				# /ST is mandatory for /SC ONCE (without it the create always failed before v2.7.0);
				# the task is started right away with /run, the time only satisfies schtasks
				$createArgs += @('/sc', 'once', '/st', (Get-Date).AddMinutes(1).ToString('HH:mm'))
			}
			$createOutput = & schtasks.exe @createArgs 2>&1
			if ($LASTEXITCODE -eq 0)
			{
				if ($IsScheduling)
				{
					Write_Log "Scheduled time sync every $IntervalMinutes minutes on lane $laneNum." "green"
					$success = $true
				}
				else
				{
					$runOutput = & schtasks.exe /run /s $machine /tn $taskName 2>&1
					if ($LASTEXITCODE -eq 0)
					{
						Write_Log "Executed time sync on lane $laneNum." "green"
						$success = $true
						# one-shot: remove the task once it had time to run
						Start-Sleep -Seconds 5
						$deleteOutput = & schtasks.exe /delete /s $machine /tn $taskName /f 2>&1
						if ($LASTEXITCODE -ne 0) { Write_Log "Delete output for [$machine]: $deleteOutput" "yellow" }
					}
					else { Write_Log "Run output for [$machine]: $runOutput" "yellow" }
				}
			}
			else { Write_Log "Create output for [$machine]: $createOutput" "yellow" }
		}
		catch
		{
			Write_Log "Exception during schtasks create/run: $_" "yellow"
		}
		# -------- SECONDARY: FILE-DROP WITH @EXEC --------
		if (-not $success)
		{
			try
			{
				$TempDir = $env:TEMP
				$xfFolder = Join-Path $OfficePath ("XF{0}{1}" -f $StoreNumber, $laneNum)

				if (-not (Test-Path $xfFolder))
				{
					Write_Log "XF folder for lane $laneNum does not exist: $xfFolder" "red"
					continue
				}

				# Always-available lane path
				$remoteShareRoot = "\\$machine\Storeman\Temp"
				if (-not (Test-Path $remoteShareRoot))
				{
					Write_Log "Lane share not reachable: $remoteShareRoot" "yellow"
					throw "Storeman\Temp not reachable"
				}

				$remoteBatPath = Join-Path $remoteShareRoot "sync_time.bat"

				# Create local BAT (ASCII)
				$batContent = "@echo off`r`nnet time \\$serverIP /set /yes"
				$localBatPath = Join-Path $TempDir ("sync_time_{0}.bat" -f $laneNum)
				Set-Content -Path $localBatPath -Value $batContent -Encoding Ascii

				# Copy to lane
				Copy-Item -Path $localBatPath -Destination $remoteBatPath -Force -ErrorAction Stop
				Remove-Item -Path $localBatPath -Force -ErrorAction SilentlyContinue
				Write_Log "Copied sync_time.bat to \\$machine\Storeman\Temp for lane $laneNum." "green"

				# Drop exec trigger in XF (ASCII)
				$execFilePath = Join-Path $xfFolder "exec_time_sync.txt"
				$execContent = "@EXEC(Run='" + (Join-Path $script:PrimaryStoremanRoot 'Temp\sync_time.bat') + "')"
				Write_Sms_File -Path $execFilePath -Content $execContent

				Write_Log "Created @EXEC trigger in XF folder for lane $laneNum (Storeman\Temp path)." "green"

				# The .bat stays in Storeman\Temp (overwritten next time): deleting it after 10 s
				# removed it before a busy lane had run the trigger, and the sync silently did nothing
				$success = $true
			}
			catch
			{
				Write_Log "File/@EXEC fallback failed for lane ${laneNum}: $_" "red"
			}
		}

		# -------- TERTIARY: MAILSLOT --------
		if (-not $success)
		{
			try
			{
				$mailslot = "\\$machine\mailslot\SMSStart_${StoreNumber}${laneNum}"
				$result = [MailslotSender]::SendMailslotCommand($mailslot, $wizCommand)
				if ($result)
				{
					Write_Log "Time sync sent to $machine (Lane $laneNum) via mailslot (final fallback)." "green"
					$success = $true
				}
				else
				{
					Write_Log "Mailslot failed for $machine (Lane $laneNum) [final fallback]." "yellow"
				}
			}
			catch
			{
				Write_Log "Exception during mailslot send (final fallback): $_" "yellow"
			}
		}

		if (-not $success)
		{
			Write_Log "All methods failed for lane $laneNum!" "red"
		}
	}

	Write_Log "`r`n==================== Ending Time Sync ====================`r`n" "blue"
}

# ===================================================================================================
#                              FUNCTION: Send_SERVER_time_to_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Time-sync wizard: shows the lane selection form, then hands the per-lane work to
#   Send_TIME_Sync_Lanes_Worker, running it as a background job via
#   Start_Background_Function_Job (inline fallback if the job cannot start).
#
# Parameters:
#   - StoreNumber  [string]: store to process, 3 digits (required)
#   - Schedule     [switch]: prompt for an interval (minutes) for repeated syncs
#
# Variables:
#   - $script:PrimaryStoremanRoot: storeman root (seeded to the worker job)
#   - $OfficePath: office path (seeded to the worker job)
#
# Notes:
#   - Per-lane work lives in Send_TIME_Sync_Lanes_Worker (v2.6.9 split for background jobs).
# ===================================================================================================

function Send_SERVER_time_to_Lanes
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[switch]$Schedule # Optional: enables scheduling mode (prompt for interval)
	)

	Write_Log "`r`n==================== Starting Time Sync ====================`r`n" "blue"

	# 1. Unified Lane Picker
	$laneSelection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
	if (-not $laneSelection -or -not $laneSelection.Lanes -or $laneSelection.Lanes.Count -eq 0)
	{
		Write_Log "No lanes selected for time sync." "yellow"
		Write_Log "`r`n==================== Time Sync Aborted ====================`r`n" "blue"
		return
	}
	$selectedLanes = $laneSelection.Lanes

	# 2. Scheduling Mode Prompt (Optional)
	$isScheduling = $false
	$intervalMinutes = 0
	if ($Schedule)
	{
		Add-Type -AssemblyName System.Windows.Forms
		$formInterval = New-Object System.Windows.Forms.Form
		$formInterval.Text = "Set Sync Interval"
		$formInterval.Size = New-Object System.Drawing.Size(300, 150)
		$formInterval.StartPosition = "CenterScreen"
		$label = New-Object System.Windows.Forms.Label
		$label.Text = "Enter interval in minutes:"
		$label.Location = New-Object System.Drawing.Point(20, 20)
		$formInterval.Controls.Add($label)
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Location = New-Object System.Drawing.Point(20, 50)
		$textBox.Size = New-Object System.Drawing.Size(240, 20)
		$formInterval.Controls.Add($textBox)
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Text = "OK"
		$okButton.Location = New-Object System.Drawing.Point(100, 80)
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$formInterval.Controls.Add($okButton)
		$result = $formInterval.ShowDialog()
		if ($result -ne [System.Windows.Forms.DialogResult]::OK)
		{
			Write_Log "Scheduling canceled." "yellow"
			return
		}
		if (-not [int]::TryParse($textBox.Text, [ref]$intervalMinutes) -or $intervalMinutes -le 0)
		{
			Write_Log "Invalid interval. Must be a positive integer." "red"
			return
		}
		$isScheduling = $true
	}

	# 3. (v2.6.9) the per-lane work runs as a BACKGROUND job so the
	#    UI thread never blocks. Per-lane details go to the file log as
	#    the job progresses; the shared job reaper logs the one-line
	#    completion notice to the log window.
	$bgJob = Start_Background_Function_Job `
		-FunctionName 'Send_TIME_Sync_Lanes_Worker' `
		-Label ('Time sync to ' + @($selectedLanes).Count + ' lane(s)') `
		-CallArgs @{ StoreNumber = $StoreNumber; SelectedLanes = $selectedLanes; IsScheduling = $isScheduling; IntervalMinutes = $intervalMinutes } `
		-SeedVars @{ OfficePath = $OfficePath; PrimaryStoremanRoot = $script:PrimaryStoremanRoot } `
		-DependentBodies @{
			Ensure_MailslotSender_Type = (Get-Command Ensure_MailslotSender_Type -CommandType Function).ScriptBlock.ToString()
			Write_Sms_File             = (Get-Command Write_Sms_File -CommandType Function).ScriptBlock.ToString()
		} `
		-NotifyOnComplete
	if ($bgJob)
	{
		Write_Log ("Time sync to {0} lane(s) is running in the background - per-lane details go to the file log." -f @($selectedLanes).Count) "green"
	}
	else
	{
		Write_Log "Could not start the background time-sync job - running inline." "yellow"
		Send_TIME_Sync_Lanes_Worker -StoreNumber $StoreNumber -SelectedLanes $selectedLanes -IsScheduling:$isScheduling -IntervalMinutes $intervalMinutes
	}
}

# ===================================================================================================
#                              FUNCTION: Send_SMS_Key_to_Lanes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Syncs the server SMS key to the lane databases. "Run now" uses the in-memory lane
#   protocols and tries Windows auth first, then SQL auth (Tecnica), then an SQI file
#   fallback; the scheduled task variant reads Protocol_Results.txt at runtime and uses
#   SQL auth for lane DB updates when the protocol is not File.
#
# Parameters:
#   - StoreNumber           [string]: store to process (required)
#   - LaneNumbers           [string[]]: lanes to target (default: all)
#   - NonInteractive        [switch]: run without prompts
#   - SqlTimeoutSeconds     [int]: SQL query timeout in seconds (default 30)
#   - ConnectTimeoutSeconds [int]: SQL connect timeout in seconds (default 1)
#
# Variables:
#   - $script:LaneProtocols: per-lane protocol cache used by "Run now"
#   - $script:FunctionResults: lane map and SQL module (input)
#
# Writes:
#   - lane databases (key sync) and a scheduled task for the recurring sync
# ===================================================================================================

function Send_SMS_Key_to_Lanes
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string[]]$LaneNumbers = $null,
		[Parameter(Mandatory = $false)]
		[switch]$NonInteractive,
		[Parameter(Mandatory = $false)]
		[int]$SqlTimeoutSeconds = 30,
		[Parameter(Mandatory = $false)]
		[int]$ConnectTimeoutSeconds = 1
	)

	Write_Log "`r`n==================== Starting Send_SMS_Key_to_Lanes ====================`r`n" "blue"

	# Normalize store number (3 digits)
	$StoreNumber = "$StoreNumber".PadLeft(3, '0')

	# Ensure FunctionResults exists
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }

	# SQL AUTH creds (used for scheduled; and as fallback for Run Now if WinAuth fails).
	# Was "TB$upp0rT" in double quotes: $upp0rT expanded to nothing and the password sent was "TB".
	$SqlAuthUser = $script:BackupSqlUser
	$SqlAuthPass = $script:BackupSqlPass

	# Everything here goes through System.Data.SqlClient, so no SQL module is needed

	# Ensure cached lane map exists (fallback only if missing)
	if (-not $script:FunctionResults['LaneNumToMachineName'])
	{
		Write_Log "LaneNumToMachineName missing; calling Retrieve_Nodes as fallback." "yellow"
		$null = Retrieve_Nodes -StoreNumber $StoreNumber
	}
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	if (-not $LaneNumToMachineName)
	{
		Write_Log "LaneNumToMachineName still missing; cannot continue." "red"
		Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
		return
	}

	# OfficePath required for SQI fallback
	if (-not $OfficePath -or -not (Test-Path $OfficePath))
	{
		$guessOffice = $script:StoremanOfficeRoots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
		if (-not $guessOffice) { $guessOffice = $script:PrimaryStoremanOfficePath }
		if (Test-Path $guessOffice) { $OfficePath = $guessOffice }
	}
	if (-not $OfficePath -or -not (Test-Path $OfficePath))
	{
		Write_Log "OfficePath missing/invalid; SQI fallback would fail. OfficePath='$OfficePath'." "red"
		Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
		return
	}

	# IMPORTANT:
	# - RUN NOW uses ONLY this variable. DEFAULT = File unless explicitly TCP/Named Pipes.
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }

	# Protocol results file path (Scheduled Task will read this at runtime)
	$protoFile = $script:ProtocolResultsFile

	# MODE selection (Interactive only): Run Now vs Schedule Only
	$mode = "run"
	if (-not $NonInteractive)
	{
		try { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
		catch { }

		$script:__modePick = $null
		$formMode = New-Object System.Windows.Forms.Form
		$formMode.Text = "Server Key Sync"
		$formMode.Size = New-Object System.Drawing.Size(540, 230)
		$formMode.StartPosition = "CenterScreen"
		$formMode.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$formMode.MaximizeBox = $false
		$formMode.MinimizeBox = $false

		$lblMode = New-Object System.Windows.Forms.Label
		$lblMode.Text = "Choose what you want to do:"
		$lblMode.AutoSize = $true
		$lblMode.Location = New-Object System.Drawing.Point(18, 18)
		$formMode.Controls.Add($lblMode)

		$btnRunNow = New-Object System.Windows.Forms.Button
		$btnRunNow.Text = "Run Now"
		$btnRunNow.Size = New-Object System.Drawing.Size(150, 45)
		$btnRunNow.Location = New-Object System.Drawing.Point(25, 65)
		$btnRunNow.Add_Click({ $script:__modePick = "run"; $formMode.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formMode.Close() })
		$formMode.Controls.Add($btnRunNow)

		$btnScheduleOnly = New-Object System.Windows.Forms.Button
		$btnScheduleOnly.Text = "Schedule Task Only"
		$btnScheduleOnly.Size = New-Object System.Drawing.Size(190, 45)
		$btnScheduleOnly.Location = New-Object System.Drawing.Point(190, 65)
		$btnScheduleOnly.Add_Click({ $script:__modePick = "schedule"; $formMode.DialogResult = [System.Windows.Forms.DialogResult]::OK; $formMode.Close() })
		$formMode.Controls.Add($btnScheduleOnly)

		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Text = "Cancel"
		$btnCancel.Size = New-Object System.Drawing.Size(110, 45)
		$btnCancel.Location = New-Object System.Drawing.Point(395, 65)
		$btnCancel.Add_Click({ $script:__modePick = "cancel"; $formMode.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $formMode.Close() })
		$formMode.Controls.Add($btnCancel)

		[void]$formMode.ShowDialog()
		$mode = $script:__modePick
		Remove-Variable -Name __modePick -Scope Script -ErrorAction SilentlyContinue

		if ($mode -eq "cancel" -or [string]::IsNullOrWhiteSpace($mode))
		{
			Write_Log "User cancelled Server Key Sync." "yellow"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}
	}

	# Lane selection (interactive) or normalize lanes (noninteractive)
	if (-not $NonInteractive)
	{
		Write_Log "Opening lane selection..." "yellow"
		$selection = Show_Node_Selection_Form -StoreNumber $StoreNumber -NodeTypes "Lane"
		if (-not $selection -or -not $selection.Lanes -or $selection.Lanes.Count -eq 0)
		{
			Write_Log "Lane selection cancelled or returned no lanes." "yellow"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}

		$tmp = @()
		foreach ($ln in $selection.Lanes)
		{
			$raw = $null
			if ($ln -is [pscustomobject] -and $ln.LaneNumber) { $raw = "$($ln.LaneNumber)" }
			else { $raw = "$ln" }

			$n = ($raw -replace '[^\d]', '')
			if (-not [string]::IsNullOrWhiteSpace($n)) { $tmp += $n.PadLeft(3, '0') }
		}
		$LaneNumbers = @($tmp | Where-Object { $_ } | Select-Object -Unique)
	}
	else
	{
		if (-not $LaneNumbers -or $LaneNumbers.Count -eq 0)
		{
			Write_Log "NonInteractive run requires -LaneNumbers." "red"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}
		$tmp = @()
		foreach ($ln in $LaneNumbers)
		{
			$n = ("$ln" -replace '[^\d]', '')
			if (-not [string]::IsNullOrWhiteSpace($n)) { $tmp += $n.PadLeft(3, '0') }
		}
		$LaneNumbers = @($tmp | Where-Object { $_ } | Select-Object -Unique)
	}

	Write_Log ("Selected lanes: " + ($LaneNumbers -join ", ")) "cyan"

	# Decide whether to deploy now
	$doDeployNow = $true
	if (-not $NonInteractive -and $mode -eq "schedule") { $doDeployNow = $false }

	# Results containers
	$UpdatedViaSql = New-Object System.Collections.Generic.List[string]
	$UpdatedViaFile = New-Object System.Collections.Generic.List[string]
	$FailedLanes = New-Object System.Collections.Generic.List[string]

	# =================================================================================================
	# DEPLOY NOW
	#   - Uses ONLY $script:LaneProtocols (already populated)
	#   - DEFAULT = File unless explicitly TCP/Named Pipes
	#   - Tries WinAuth first, then SQL Auth, then SQI
	# =================================================================================================
	if ($doDeployNow)
	{
		try { Add-Type -AssemblyName System.Data 2>$null }
		catch { }

		# LOCAL key read
		$localConnStr = $script:FunctionResults['ConnectionString']
		if ([string]::IsNullOrWhiteSpace($localConnStr) -or $localConnStr -eq "N/A")
		{
			Write_Log "Local DB ConnectionString missing (FunctionResults['ConnectionString'])." "red"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}

		# Enforce connect timeout on LOCAL conn string
		$localConnStrEff = $localConnStr
		if ($localConnStrEff -notmatch '(?i)\bConnect\s*Timeout\s*=')
		{
			if ($localConnStrEff.Trim().EndsWith(';')) { $localConnStrEff += "Connect Timeout=$ConnectTimeoutSeconds;" }
			else { $localConnStrEff += ";Connect Timeout=$ConnectTimeoutSeconds;" }
		}
		else
		{
			$localConnStrEff = [regex]::Replace($localConnStrEff, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', "Connect Timeout=$ConnectTimeoutSeconds")
		}

		Write_Log "Reading SYS_TAB.F1243 from LOCAL DB..." "yellow"

		$rawValues = @()
		try
		{
			$vals = New-Object System.Collections.Generic.List[string]
			$cn = New-Object System.Data.SqlClient.SqlConnection $localConnStrEff
			try
			{
				$cn.Open()
				$cmd = $cn.CreateCommand()
				$cmd.CommandText = "SELECT F1243 FROM SYS_TAB;"
				$cmd.CommandTimeout = $SqlTimeoutSeconds
				$rdr = $cmd.ExecuteReader()
				try
				{
					while ($rdr.Read())
					{
						if (-not $rdr.IsDBNull(0)) { $vals.Add([string]$rdr.GetValue(0)) | Out-Null }
					}
				}
				finally
				{
					try { $rdr.Close() }
					catch { }
					try { $rdr.Dispose() }
					catch { }
				}
			}
			finally
			{
				try { $cn.Close() }
				catch { }
				try { $cn.Dispose() }
				catch { }
			}
			$rawValues = @($vals)
		}
		catch { $rawValues = @() }

		$serverKey = $null
		$bestLen = -1
		foreach ($v in $rawValues)
		{
			if ($v -eq $null) { continue }
			$t = ([string]$v).Trim()
			$stripped = [regex]::Replace($t, '[\s\u00A0]+', '')
			if (-not [string]::IsNullOrWhiteSpace($t) -and -not [string]::IsNullOrWhiteSpace($stripped))
			{
				if ($t.Length -gt $bestLen) { $bestLen = $t.Length; $serverKey = $t }
			}
		}

		if ([string]::IsNullOrWhiteSpace($serverKey))
		{
			Write_Log "Server key (SYS_TAB.F1243) came back empty/null from LOCAL DB. Aborting." "red"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}

		Write_Log ("Server key retrieved (length {0})." -f $serverKey.Length) "green"

		$safeServerKey = $serverKey.Replace("'", "''")
		$updateSql = "UPDATE SYS_TAB SET F1243 = '$safeServerKey'"

		# SQI fallback file (ANSI 1252, no BOM)
		$SqiFileName = "UPDATE_SMS_KEY.SQI" # same name as the scheduled task, so a lane never queues two
		$SqiContent = "@dbEXEC($updateSql)`r`n"
		$ansiEncoding = [System.Text.Encoding]::GetEncoding(1252)
		$TempSqiPath = Join-Path ([System.IO.Path]::GetTempPath()) $SqiFileName

		try
		{
			[System.IO.File]::WriteAllText($TempSqiPath, $SqiContent, $ansiEncoding)
			try { attrib -a -r "$TempSqiPath" 2>$null | Out-Null }
			catch { }
			try { Set-ItemProperty -Path $TempSqiPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal) -ErrorAction SilentlyContinue }
			catch { }
		}
		catch
		{
			Write_Log "Failed to create temp SQI '$TempSqiPath'. Error: $_" "red"
			Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
			return
		}

		foreach ($LaneNumber in $LaneNumbers)
		{
			Write_Log "`r`n--- Lane $LaneNumber ---" "magenta"

			$laneInfo = $null
			try { $laneInfo = Get_All_Lanes_Database_Info -LaneNumber $LaneNumber }
			catch { $laneInfo = $null }
			if (-not $laneInfo)
			{
				Write_Log "Could not retrieve cached DB info for lane $LaneNumber. Skipping." "red"
				$FailedLanes.Add($LaneNumber) | Out-Null
				continue
			}

			# Normalize PSCustomObject -> hashtable (so ['TcpConnStr'] works)
			if ($laneInfo -isnot [hashtable])
			{
				$h = @{ }
				foreach ($p in $laneInfo.PSObject.Properties) { $h[$p.Name] = $p.Value }
				$laneInfo = $h
			}

			$tcpConnStr = $laneInfo['TcpConnStr']
			$namedPipesConnStr = $laneInfo['NamedPipesConnStr']

			# RUN NOW: protocol comes ONLY from $script:LaneProtocols; DEFAULT = File
			$machineName = $LaneNumToMachineName[$LaneNumber]
			$protocolPref = $null
			if ($script:LaneProtocols.ContainsKey($LaneNumber)) { $protocolPref = [string]$script:LaneProtocols[$LaneNumber] }
			elseif ($machineName -and $script:LaneProtocols.ContainsKey($machineName)) { $protocolPref = [string]$script:LaneProtocols[$machineName] }
			elseif ($machineName -and $script:LaneProtocols.ContainsKey($machineName.ToLower())) { $protocolPref = [string]$script:LaneProtocols[$machineName.ToLower()] }

			$protocolPrefNorm = "File"
			if (-not [string]::IsNullOrWhiteSpace($protocolPref))
			{
				$t = $protocolPref.Trim()
				if ($t -match '(?i)^tcp$') { $protocolPrefNorm = "TCP" }
				elseif ($t -match '(?i)pipe') { $protocolPrefNorm = "Named Pipes" }
				elseif ($t -match '(?i)^offline$') { $protocolPrefNorm = "Offline" }
				else { $protocolPrefNorm = "File" }
			}

			if ($protocolPrefNorm -eq "Offline")
			{
				Write_Log "Lane $LaneNumber ($machineName) is offline. Skipping." "yellow"
				$FailedLanes.Add($LaneNumber) | Out-Null
				continue
			}

			$sqlWorked = $false

			# SQL attempts ONLY when protocol explicitly says TCP/Named Pipes
			if ($protocolPrefNorm -ne "File")
			{
				$tryOrder = @("TCP", "Named Pipes")
				if ($protocolPrefNorm -eq "Named Pipes") { $tryOrder = @("Named Pipes", "TCP") }

				foreach ($proto in $tryOrder)
				{
					$connToUse = if ($proto -eq "Named Pipes") { $namedPipesConnStr }
					else { $tcpConnStr }
					if ([string]::IsNullOrWhiteSpace($connToUse)) { continue }

					# -----------------------
					# Attempt 1: Windows Auth
					# -----------------------
					$connWin = $connToUse
					$connWin = [regex]::Replace($connWin, '(?i)\bUser\s*ID\s*=\s*[^;]*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bUID\s*=\s*[^;]*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bPassword\s*=\s*[^;]*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bPWD\s*=\s*[^;]*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bTrusted_Connection\s*=\s*Yes\s*;?', '')
					$connWin = [regex]::Replace($connWin, '(?i)\bIntegrated\s+Security\s*=\s*False\s*;?', '')

					if ($connWin -notmatch '(?i)\bIntegrated\s+Security\s*=')
					{
						if ($connWin.Trim().EndsWith(';')) { $connWin += "Integrated Security=True;" }
						else { $connWin += ";Integrated Security=True;" }
					}
					else
					{
						$connWin = [regex]::Replace($connWin, '(?i)\bIntegrated\s+Security\s*=\s*[^;]*', 'Integrated Security=True')
					}

					if ($connWin -notmatch '(?i)\bConnect\s*Timeout\s*=')
					{
						if ($connWin.Trim().EndsWith(';')) { $connWin += "Connect Timeout=$ConnectTimeoutSeconds;" }
						else { $connWin += ";Connect Timeout=$ConnectTimeoutSeconds;" }
					}
					else
					{
						$connWin = [regex]::Replace($connWin, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', "Connect Timeout=$ConnectTimeoutSeconds")
					}

					try
					{
						$cn2 = New-Object System.Data.SqlClient.SqlConnection $connWin
						try
						{
							$cn2.Open()
							$cmd2 = $cn2.CreateCommand()
							$cmd2.CommandText = $updateSql
							$cmd2.CommandTimeout = $SqlTimeoutSeconds
							[void]$cmd2.ExecuteNonQuery()
							$sqlWorked = $true
							Write_Log "Updated via SQL ($proto) on lane $LaneNumber (Windows Auth)." "green"
							$UpdatedViaSql.Add($LaneNumber) | Out-Null
							break
						}
						finally
						{
							try { $cn2.Close() }
							catch { }
							try { $cn2.Dispose() }
							catch { }
						}
					}
					catch
					{
						$sqlWorked = $false
					}

					if ($sqlWorked) { break }

					# -----------------------
					# Attempt 2: SQL Auth
					# -----------------------
					$connSql = $connToUse
					$connSql = [regex]::Replace($connSql, '(?i)\bIntegrated\s+Security\s*=\s*True\s*;?', '')
					$connSql = [regex]::Replace($connSql, '(?i)\bTrusted_Connection\s*=\s*Yes\s*;?', '')

					if ($connSql -notmatch '(?i)\bUser\s*ID\s*=')
					{
						if ($connSql.Trim().EndsWith(';')) { $connSql += "User ID=$SqlAuthUser;" }
						else { $connSql += ";User ID=$SqlAuthUser;" }
					}
					if ($connSql -notmatch '(?i)\bPassword\s*=')
					{
						if ($connSql.Trim().EndsWith(';')) { $connSql += "Password=$SqlAuthPass;" }
						else { $connSql += ";Password=$SqlAuthPass;" }
					}

					if ($connSql -notmatch '(?i)\bConnect\s*Timeout\s*=')
					{
						if ($connSql.Trim().EndsWith(';')) { $connSql += "Connect Timeout=$ConnectTimeoutSeconds;" }
						else { $connSql += ";Connect Timeout=$ConnectTimeoutSeconds;" }
					}
					else
					{
						$connSql = [regex]::Replace($connSql, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', "Connect Timeout=$ConnectTimeoutSeconds")
					}

					try
					{
						$cn3 = New-Object System.Data.SqlClient.SqlConnection $connSql
						try
						{
							$cn3.Open()
							$cmd3 = $cn3.CreateCommand()
							$cmd3.CommandText = $updateSql
							$cmd3.CommandTimeout = $SqlTimeoutSeconds
							[void]$cmd3.ExecuteNonQuery()
							$sqlWorked = $true
							Write_Log "Updated via SQL ($proto) on lane $LaneNumber (SQL Auth)." "green"
							$UpdatedViaSql.Add($LaneNumber) | Out-Null
							break
						}
						finally
						{
							try { $cn3.Close() }
							catch { }
							try { $cn3.Dispose() }
							catch { }
						}
					}
					catch
					{
						$sqlWorked = $false
					}

					if ($sqlWorked) { break }
				}
			}

			# SQI fallback (protocol File OR SQL failed)
			if (-not $sqlWorked)
			{
				$LaneXFDir = Join-Path $OfficePath ("XF{0}{1}" -f $StoreNumber, $LaneNumber)
				$DestSqiPath = Join-Path $LaneXFDir $SqiFileName

				if (-not (Test-Path $LaneXFDir))
				{
					Write_Log "Lane XF folder not found: $LaneXFDir" "red"
					$FailedLanes.Add($LaneNumber) | Out-Null
					continue
				}

				try
				{
					Copy-Item -Path $TempSqiPath -Destination $DestSqiPath -Force
					try { attrib -a -r "$DestSqiPath" 2>$null | Out-Null }
					catch { }
					try { Set-ItemProperty -Path $DestSqiPath -Name Attributes -Value ([System.IO.FileAttributes]::Normal) -ErrorAction SilentlyContinue }
					catch { }

					$UpdatedViaFile.Add($LaneNumber) | Out-Null
					Write_Log "Copied SQI to $DestSqiPath (archive cleared)" "green"
				}
				catch
				{
					Write_Log "Failed to copy SQI for lane $LaneNumber. Error: $_" "red"
					$FailedLanes.Add($LaneNumber) | Out-Null
				}
			}
		}

		if ($UpdatedViaSql.Count -gt 0) { Write_Log ("Updated via SQL: " + ($UpdatedViaSql -join ", ")) "green" }
		if ($UpdatedViaFile.Count -gt 0) { Write_Log ("Updated via SQI: " + ($UpdatedViaFile -join ", ")) "green" }
		if ($FailedLanes.Count -gt 0) { Write_Log ("Failed lanes: " + ($FailedLanes -join ", ")) "red" }
	}

	# Restart prompt (interactive Run Now only)
	if (-not $NonInteractive -and $mode -eq "run" -and $doDeployNow)
	{
		$successful = @(@($UpdatedViaSql) + @($UpdatedViaFile) | Where-Object { $_ } | Select-Object -Unique)
		if ($successful.Count -gt 0)
		{
			try { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
			catch { }
			$ansRestart = [System.Windows.Forms.MessageBox]::Show(
				"Server key deployed successfully to: " + ($successful -join ", ") + "`r`n`r`nRestart ALL programs on these lanes now?",
				"Restart Programs?",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question
			)

			if ($ansRestart -eq [System.Windows.Forms.DialogResult]::Yes)
			{
				foreach ($ln in $successful)
				{
					$machineName = $LaneNumToMachineName[$ln]
					if ([string]::IsNullOrWhiteSpace($machineName))
					{
						Write_Log "Could not resolve machine name for lane $ln (restart skipped)." "yellow"
						continue
					}

					try
					{
						$mailslotAddress = "\\$machineName\mailslot\SMSStart_${StoreNumber}${ln}"
						$sent = [MailslotSender]::SendMailslotCommand($mailslotAddress, "@exec(RESTART_ALL=PROGRAMS).")
						if ($sent) { Write_Log "Restart sent to $machineName (lane $ln)." "green" }
						else { Write_Log "Restart FAILED to $machineName (lane $ln)." "red" }
					}
					catch
					{
						Write_Log "Restart attempt failed for lane $ln. Error: $_" "red"
					}
				}
			}
			else
			{
				Write_Log "User chose NOT to restart programs after deploy." "yellow"
			}
		}
	}

	# =================================================================================================
	# SCHEDULE (interactive only) -> ONE self-sufficient PS1
	#   - Uses SQL AUTH for lane DB updates when protocol != File
	#   - Reads Protocol_Results.txt at runtime (DEFAULT File unless explicitly TCP/Named Pipes)
	# =================================================================================================
	if (-not $NonInteractive)
	{
		$doSchedule = $false
		if ($mode -eq "schedule") { $doSchedule = $true }
		else
		{
			try { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
			catch { }
			$ansSchedule = [System.Windows.Forms.MessageBox]::Show(
				"Do you want to create/update a Scheduled Task for the selected lanes?",
				"Create Scheduled Task?",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question
			)
			if ($ansSchedule -eq [System.Windows.Forms.DialogResult]::Yes) { $doSchedule = $true }
		}

		if ($doSchedule)
		{
			try { Add-Type -AssemblyName System.Windows.Forms | Out-Null }
			catch { }

			$formFreq = New-Object System.Windows.Forms.Form
			$formFreq.Text = "Server Key Sync Frequency"
			$formFreq.Width = 520
			$formFreq.Height = 240
			$formFreq.StartPosition = "CenterScreen"

			$lbl = New-Object System.Windows.Forms.Label
			$lbl.AutoSize = $true
			$lbl.Location = New-Object System.Drawing.Point(15, 20)
			$lbl.Text = "Run every how many minutes? (1 - 1439)"
			$formFreq.Controls.Add($lbl)

			$tb = New-Object System.Windows.Forms.TextBox
			$tb.Location = New-Object System.Drawing.Point(20, 50)
			$tb.Width = 100
			$tb.Text = "60"
			$formFreq.Controls.Add($tb)

			$lbl2 = New-Object System.Windows.Forms.Label
			$lbl2.AutoSize = $true
			$lbl2.Location = New-Object System.Drawing.Point(15, 95)
			$lbl2.Text = "Selected lanes: " + ($LaneNumbers -join ", ")
			$formFreq.Controls.Add($lbl2)

			$okBtn = New-Object System.Windows.Forms.Button
			$okBtn.Text = "OK"
			$okBtn.Location = New-Object System.Drawing.Point(300, 160)
			$okBtn.Add_Click({ $formFreq.DialogResult = [System.Windows.Forms.DialogResult]::OK })
			$formFreq.Controls.Add($okBtn)

			$cancelBtn = New-Object System.Windows.Forms.Button
			$cancelBtn.Text = "Cancel"
			$cancelBtn.Location = New-Object System.Drawing.Point(380, 160)
			$cancelBtn.Add_Click({ $formFreq.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
			$formFreq.Controls.Add($cancelBtn)

			$formFreq.AcceptButton = $okBtn
			$formFreq.CancelButton = $cancelBtn

			if ($formFreq.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK)
			{
				Write_Log "Scheduled Task setup cancelled by user." "yellow"
				Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
				return
			}

			[int]$freqMinutes = 0
			if (-not [int]::TryParse($tb.Text, [ref]$freqMinutes) -or $freqMinutes -lt 1 -or $freqMinutes -gt 1439)
			{
				Write_Log "Invalid frequency entered. Using 60 minutes." "yellow"
				$freqMinutes = 60
			}

			# Required cached data for self-sufficient script
			$localConnStr = $script:FunctionResults['ConnectionString']
			if ([string]::IsNullOrWhiteSpace($localConnStr) -or $localConnStr -eq "N/A")
			{
				Write_Log "Cannot schedule because local ConnectionString is missing in FunctionResults['ConnectionString']." "red"
				Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
				return
			}

			$laneConnMap = @{ }
			foreach ($ln in $LaneNumbers)
			{
				$li = $null
				try { $li = Get_All_Lanes_Database_Info -LaneNumber $ln }
				catch { $li = $null }
				if ($li)
				{
					if ($li -isnot [hashtable])
					{
						$h2 = @{ }
						foreach ($p in $li.PSObject.Properties) { $h2[$p.Name] = $p.Value }
						$li = $h2
					}

					$laneConnMap[$ln] = @{
						TcpConnStr	      = $li['TcpConnStr']
						NamedPipesConnStr = $li['NamedPipesConnStr']
					}
				}
			}

			$taskScriptPath = Join-Path $script:ScriptsFolder ("Send_SMS_Key_Task_Store{0}.ps1" -f $StoreNumber)
			$latestLogPath = Join-Path $script:ScriptsLog    ("Send_SMS_Key_LATEST_Store{0}.txt" -f $StoreNumber)

			$enc = [System.Text.Encoding]::UTF8
			$localConnB64 = [Convert]::ToBase64String($enc.GetBytes($localConnStr))
			$laneConnB64 = [Convert]::ToBase64String($enc.GetBytes(($laneConnMap | ConvertTo-Json -Depth 6)))

			$laneListLiteral = ($LaneNumbers | ForEach-Object { "'$_'" }) -join ", "

			$sb = New-Object System.Text.StringBuilder
			[void]$sb.AppendLine("# AUTO-GENERATED Send_SMS_Key task (self-sufficient, one file)")
			[void]$sb.AppendLine("`$ErrorActionPreference = 'Stop'")
			[void]$sb.AppendLine("try { Add-Type -AssemblyName System.Data 2>`$null } catch { }")
			[void]$sb.AppendLine("`$SqlUser = '" + $SqlAuthUser.Replace("'", "''") + "'")
			[void]$sb.AppendLine("`$SqlPass = '" + $SqlAuthPass.Replace("'", "''") + "'")
			[void]$sb.AppendLine("`$StoreNumber = '$StoreNumber'")
			[void]$sb.AppendLine("`$OfficePath  = '$OfficePath'")
			[void]$sb.AppendLine("`$SqlTimeoutSeconds = $SqlTimeoutSeconds")
			[void]$sb.AppendLine("`$ConnectTimeoutSeconds = $ConnectTimeoutSeconds")
			[void]$sb.AppendLine("`$LatestLogPath = '$latestLogPath'")
			[void]$sb.AppendLine("`$ProtocolFile = '$protoFile'")
			[void]$sb.AppendLine("`$enc = [System.Text.Encoding]::UTF8")
			[void]$sb.AppendLine("`$LocalConnStr = `$enc.GetString([Convert]::FromBase64String('$localConnB64'))")
			[void]$sb.AppendLine("`$LaneConnMap  = ConvertFrom-Json (`$enc.GetString([Convert]::FromBase64String('$laneConnB64')))")
			[void]$sb.AppendLine("`$LaneNumbers = @($laneListLiteral)")
			[void]$sb.AppendLine("`$UpdatedViaSql  = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("`$UpdatedViaFile = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("`$FailedLanes    = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("# Load protocol map from txt (DEFAULT = File unless explicitly TCP/Named Pipes)")
			[void]$sb.AppendLine("`$ProtoMap = @{}")
			[void]$sb.AppendLine("if (Test-Path `$ProtocolFile) {")
			[void]$sb.AppendLine("  try {")
			[void]$sb.AppendLine("    foreach (`$line in (Get-Content -LiteralPath `$ProtocolFile -ErrorAction SilentlyContinue)) {")
			[void]$sb.AppendLine("      if (`$line -match '^\s*([^,]+),\s*([^,]+)\s*$') {")
			[void]$sb.AppendLine("        `$laneRaw = `$matches[1].Trim()")
			[void]$sb.AppendLine("        `$protoRaw = `$matches[2].Trim()")
			[void]$sb.AppendLine("        `$num = (`$laneRaw -replace '[^\d]','').PadLeft(3,'0')")
			[void]$sb.AppendLine("        if (-not [string]::IsNullOrWhiteSpace(`$num)) { `$ProtoMap[`$num] = `$protoRaw }")
			[void]$sb.AppendLine("      }")
			[void]$sb.AppendLine("    }")
			[void]$sb.AppendLine("  } catch { }")
			[void]$sb.AppendLine("}")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("# Enforce connect timeout on LOCAL conn string")
			[void]$sb.AppendLine("`$localEff = `$LocalConnStr")
			[void]$sb.AppendLine("if (`$localEff -notmatch '(?i)\bConnect\s*Timeout\s*=') { if (`$localEff.Trim().EndsWith(';')) { `$localEff += ""Connect Timeout=`$ConnectTimeoutSeconds;"" } else { `$localEff += "";Connect Timeout=`$ConnectTimeoutSeconds;"" } }")
			[void]$sb.AppendLine("else { `$localEff = [regex]::Replace(`$localEff, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', ""Connect Timeout=`$ConnectTimeoutSeconds"") }")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("# Read key")
			[void]$sb.AppendLine("`$rawValues = @()")
			[void]$sb.AppendLine("try {")
			[void]$sb.AppendLine("  `$vals = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("  `$cn = New-Object System.Data.SqlClient.SqlConnection `$localEff")
			[void]$sb.AppendLine("  try { `$cn.Open(); `$cmd = `$cn.CreateCommand(); `$cmd.CommandText = ""SELECT F1243 FROM SYS_TAB;""; `$cmd.CommandTimeout = `$SqlTimeoutSeconds; `$rdr = `$cmd.ExecuteReader();")
			[void]$sb.AppendLine("    try { while (`$rdr.Read()) { if (-not `$rdr.IsDBNull(0)) { `$vals.Add([string]`$rdr.GetValue(0)) | Out-Null } } } finally { try { `$rdr.Close() } catch { } try { `$rdr.Dispose() } catch { } }")
			[void]$sb.AppendLine("  } finally { try { `$cn.Close() } catch { } try { `$cn.Dispose() } catch { } }")
			[void]$sb.AppendLine("  `$rawValues = @(`$vals)")
			[void]$sb.AppendLine("} catch { `$rawValues = @() }")
			[void]$sb.AppendLine("`$serverKey = `$null; `$bestLen = -1")
			[void]$sb.AppendLine("foreach (`$v in `$rawValues) { if (`$v -eq `$null) { continue }; `$t = ([string]`$v).Trim(); `$stripped = [regex]::Replace(`$t,'[\s\u00A0]+',''); if (-not [string]::IsNullOrWhiteSpace(`$t) -and -not [string]::IsNullOrWhiteSpace(`$stripped)) { if (`$t.Length -gt `$bestLen) { `$bestLen = `$t.Length; `$serverKey = `$t } } }")
			[void]$sb.AppendLine("# SYSTEM usually has no Windows login on the store DB: retry the read with SQL auth")
			[void]$sb.AppendLine("if ([string]::IsNullOrWhiteSpace(`$serverKey)) {")
			[void]$sb.AppendLine("  `$sqlEff = [regex]::Replace(`$localEff, '(?i)\b(Integrated\s+Security|Trusted_Connection)\s*=\s*[^;]*;?', '')")
			[void]$sb.AppendLine("  `$sqlEff = `$sqlEff.TrimEnd(';') + "";User ID=`$SqlUser;Password=`$SqlPass;""")
			[void]$sb.AppendLine("  try { `$cn = New-Object System.Data.SqlClient.SqlConnection `$sqlEff; try { `$cn.Open(); `$cmd = `$cn.CreateCommand(); `$cmd.CommandText = ""SELECT TOP 1 F1243 FROM SYS_TAB WHERE LEN(LTRIM(RTRIM(F1243))) > 0 ORDER BY LEN(F1243) DESC;""; `$cmd.CommandTimeout = `$SqlTimeoutSeconds; `$v = `$cmd.ExecuteScalar(); if (`$v -ne `$null -and `$v -isnot [DBNull]) { `$serverKey = ([string]`$v).Trim() } } finally { `$cn.Dispose() } } catch { }")
			[void]$sb.AppendLine("}")
			[void]$sb.AppendLine("if ([string]::IsNullOrWhiteSpace(`$serverKey)) {")
			[void]$sb.AppendLine("  `$logDir = [System.IO.Path]::GetDirectoryName(`$LatestLogPath); if (-not (Test-Path `$logDir)) { New-Item -Path `$logDir -ItemType Directory -Force | Out-Null }")
			[void]$sb.AppendLine("  @((""Timestamp : {0}"" -f (Get-Date -Format ""yyyy-MM-dd HH:mm:ss"")), ""Status    : FAILED"", ""Reason    : server key (SYS_TAB.F1243) could not be read from the local DB"") | Set-Content -Path `$LatestLogPath -Encoding UTF8")
			[void]$sb.AppendLine("  exit 1")
			[void]$sb.AppendLine("}")
			[void]$sb.AppendLine("`$safeServerKey = `$serverKey.Replace(""'"",""''"")")
			# (only written when the lane differs: no needless row update on every run)
			[void]$sb.AppendLine("`$updateSql = ""UPDATE SYS_TAB SET F1243 = '"" + `$safeServerKey + ""' WHERE ISNULL(F1243, '') <> '"" + `$safeServerKey + ""'""")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("# SQI (ANSI/1252)")
			[void]$sb.AppendLine("`$SqiFileName = ""UPDATE_SMS_KEY.SQI""")
			[void]$sb.AppendLine("`$SqiContent  = ""@dbEXEC("" + `$updateSql + "")`r`n""")
			[void]$sb.AppendLine("`$ansi = [System.Text.Encoding]::GetEncoding(1252)")
			[void]$sb.AppendLine("`$TempSqiPath = Join-Path ([System.IO.Path]::GetTempPath()) `$SqiFileName")
			[void]$sb.AppendLine("[System.IO.File]::WriteAllText(`$TempSqiPath, `$SqiContent, `$ansi)")
			[void]$sb.AppendLine("try { attrib -a -r `"`$TempSqiPath`" 2>`$null | Out-Null } catch { }")
			# File lanes: the SQI is dropped again only when the key changed or a day has passed (it was
			# dropped on every run before v2.7.7 - the lanes reprocessed the same file all day)
			[void]$sb.AppendLine('$KeyHash = [BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash($enc.GetBytes($serverKey))).Replace("-", "")')
			[void]$sb.AppendLine('$StatePath = [System.IO.Path]::ChangeExtension($LatestLogPath, ".state.json")')
			[void]$sb.AppendLine('$State = @{}')
			[void]$sb.AppendLine('if (Test-Path -LiteralPath $StatePath) { try { $so = Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json; foreach ($sp in $so.PSObject.Properties) { $State[$sp.Name] = $sp.Value } } catch { } }')
			[void]$sb.AppendLine('$SkippedLanes = New-Object System.Collections.Generic.List[string]')
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("foreach (`$ln in `$LaneNumbers) {")
			[void]$sb.AppendLine("  `$pref = 'File'")
			[void]$sb.AppendLine("  if (`$ProtoMap.ContainsKey(`$ln)) { `$pref = [string]`$ProtoMap[`$ln] }")
			[void]$sb.AppendLine("  `$prefNorm = 'File'")
			[void]$sb.AppendLine("  if (-not [string]::IsNullOrWhiteSpace(`$pref)) { if (`$pref -match '(?i)^tcp$') { `$prefNorm='TCP' } elseif (`$pref -match '(?i)pipe') { `$prefNorm='Named Pipes' } }")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("  `$laneProp = `$LaneConnMap.PSObject.Properties[`$ln]")
			[void]$sb.AppendLine("  `$laneObj  = if (`$laneProp) { `$laneProp.Value } else { `$null }")
			[void]$sb.AppendLine("  `$tcpConn  = if (`$laneObj) { [string]`$laneObj.TcpConnStr } else { `$null }")
			[void]$sb.AppendLine("  `$npConn   = if (`$laneObj) { [string]`$laneObj.NamedPipesConnStr } else { `$null }")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine("  `$sqlWorked = `$false")
			[void]$sb.AppendLine("  if (`$prefNorm -ne 'File') {")
			[void]$sb.AppendLine("    `$tryOrder = @('TCP','Named Pipes'); if (`$prefNorm -eq 'Named Pipes') { `$tryOrder=@('Named Pipes','TCP') }")
			[void]$sb.AppendLine("    foreach (`$proto in `$tryOrder) {")
			[void]$sb.AppendLine("      `$connToUse = if (`$proto -eq 'Named Pipes') { `$npConn } else { `$tcpConn }")
			[void]$sb.AppendLine("      if ([string]::IsNullOrWhiteSpace(`$connToUse)) { continue }")
			[void]$sb.AppendLine("      `$connEff = `$connToUse")
			[void]$sb.AppendLine("      `$connEff = [regex]::Replace(`$connEff, '(?i)\bIntegrated\s+Security\s*=\s*True\s*;?', '')")
			[void]$sb.AppendLine("      `$connEff = [regex]::Replace(`$connEff, '(?i)\bTrusted_Connection\s*=\s*Yes\s*;?', '')")
			[void]$sb.AppendLine("      if (`$connEff -notmatch '(?i)\bUser\s*ID\s*=') { if (`$connEff.Trim().EndsWith(';')) { `$connEff += ""User ID=`$SqlUser;"" } else { `$connEff += "";User ID=`$SqlUser;"" } }")
			[void]$sb.AppendLine("      if (`$connEff -notmatch '(?i)\bPassword\s*=') { if (`$connEff.Trim().EndsWith(';')) { `$connEff += ""Password=`$SqlPass;"" } else { `$connEff += "";Password=`$SqlPass;"" } }")
			[void]$sb.AppendLine("      if (`$connEff -notmatch '(?i)\bConnect\s*Timeout\s*=') { if (`$connEff.Trim().EndsWith(';')) { `$connEff += ""Connect Timeout=`$ConnectTimeoutSeconds;"" } else { `$connEff += "";Connect Timeout=`$ConnectTimeoutSeconds;"" } }")
			[void]$sb.AppendLine("      else { `$connEff = [regex]::Replace(`$connEff, '(?i)\bConnect\s*Timeout\s*=\s*\d+\s*', ""Connect Timeout=`$ConnectTimeoutSeconds"") }")
			[void]$sb.AppendLine("      try { `$cn2 = New-Object System.Data.SqlClient.SqlConnection `$connEff; try { `$cn2.Open(); `$cmd2 = `$cn2.CreateCommand(); `$cmd2.CommandText = `$updateSql; `$cmd2.CommandTimeout = `$SqlTimeoutSeconds; [void]`$cmd2.ExecuteNonQuery(); `$sqlWorked = `$true } finally { try { `$cn2.Close() } catch { } try { `$cn2.Dispose() } catch { } } } catch { `$sqlWorked = `$false }")
			[void]$sb.AppendLine("      if (`$sqlWorked) { `$UpdatedViaSql.Add(`$ln) | Out-Null; break }")
			[void]$sb.AppendLine("    }")
			[void]$sb.AppendLine("  }")
			[void]$sb.AppendLine("  if (-not `$sqlWorked) {")
			[void]$sb.AppendLine('    $prev = $State[$ln]')
			[void]$sb.AppendLine('    if ($prev -and [string]$prev.Hash -eq $KeyHash -and $prev.Sent -and ((Get-Date) - [datetime]$prev.Sent).TotalHours -lt 24) { $SkippedLanes.Add($ln) | Out-Null; continue }')
			[void]$sb.AppendLine("    `$laneXF = Join-Path `$OfficePath (""XF{0}{1}"" -f `$StoreNumber, `$ln)")
			[void]$sb.AppendLine("    `$dest = Join-Path `$laneXF `$SqiFileName")
			[void]$sb.AppendLine("    if (-not (Test-Path `$laneXF)) { `$FailedLanes.Add(`$ln) | Out-Null }")
			[void]$sb.AppendLine("    else { try { Copy-Item -Path `$TempSqiPath -Destination `$dest -Force; try { attrib -a -r `"`$dest`" 2>`$null | Out-Null } catch { }; `$UpdatedViaFile.Add(`$ln) | Out-Null } catch { `$FailedLanes.Add(`$ln) | Out-Null } }")
			[void]$sb.AppendLine("  }")
			[void]$sb.AppendLine("}")
			[void]$sb.AppendLine("")
			[void]$sb.AppendLine('foreach ($done in @($UpdatedViaFile)) { $State[$done] = @{ Hash = $KeyHash; Sent = (Get-Date).ToString("o") } }')
			[void]$sb.AppendLine('try { ($State | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $StatePath -Encoding UTF8 } catch { }')
			[void]$sb.AppendLine("`$okAll = @(@(`$UpdatedViaSql) + @(`$UpdatedViaFile) | Where-Object { `$_ } | Select-Object -Unique)")
			[void]$sb.AppendLine("`$status = ""FAILED""; if (`$okAll.Count -gt 0) { `$status = ""OK""; if (`$FailedLanes.Count -gt 0) { `$status = ""PARTIAL"" } }")
			[void]$sb.AppendLine("`$okAllText = ""(none)""; if (`$okAll.Count -gt 0) { `$okAllText = (`$okAll -join "", "") }")
			[void]$sb.AppendLine("`$sqlText = ""(none)""; if (`$UpdatedViaSql.Count -gt 0) { `$sqlText = (@(`$UpdatedViaSql) -join "", "") }")
			[void]$sb.AppendLine("`$sqiText = ""(none)""; if (`$UpdatedViaFile.Count -gt 0) { `$sqiText = (@(`$UpdatedViaFile) -join "", "") }")
			[void]$sb.AppendLine("`$failedText = ""(none)""; if (`$FailedLanes.Count -gt 0) { `$failedText = (@(`$FailedLanes) -join "", "") }")
			[void]$sb.AppendLine("`$lines = New-Object System.Collections.Generic.List[string]")
			[void]$sb.AppendLine("`$lines.Add((""Timestamp : {0}"" -f (Get-Date -Format ""yyyy-MM-dd HH:mm:ss""))) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Computer  : {0}"" -f `$env:COMPUTERNAME)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Store     : {0}"" -f `$StoreNumber)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""ProtocolFile : {0}"" -f `$ProtocolFile)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Status    : {0}"" -f `$status)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Selected  : {0}"" -f (`$LaneNumbers -join "", ""))) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""OK (ALL)  : {0}"" -f `$okAllText)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""OK (SQL)  : {0}"" -f `$sqlText)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""OK (SQI)  : {0}"" -f `$sqiText)) | Out-Null")
			[void]$sb.AppendLine("`$lines.Add((""Failed    : {0}"" -f `$failedText)) | Out-Null")
			[void]$sb.AppendLine('$lines.Add(("Unchanged : {0}" -f $(if ($SkippedLanes.Count -gt 0) { $SkippedLanes -join ", " } else { "(none, or SQL lanes)" }))) | Out-Null')
			[void]$sb.AppendLine("`$logDir = [System.IO.Path]::GetDirectoryName(`$LatestLogPath)")
			[void]$sb.AppendLine("if (-not (Test-Path `$logDir)) { New-Item -Path `$logDir -ItemType Directory -Force | Out-Null }")
			[void]$sb.AppendLine("`$lines | Set-Content -Path `$LatestLogPath -Encoding UTF8")

			try
			{
				Set-Content -Path $taskScriptPath -Value ($sb.ToString() -replace "`r?`n", "`r`n") -Encoding UTF8
				Write_Log "Task script written/updated (single file): $taskScriptPath" "green"
				Write_Log "Latest results log (overwrites each run): $latestLogPath" "cyan"
			}
			catch
			{
				Write_Log "Failed to write task script '$taskScriptPath'. Error: $_" "red"
				Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
				return
			}

			try
			{
				$taskName = "Send_SMS_Key - Store $StoreNumber"
				$psExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
				# \" survives 5.1's native argument passing (plain quotes were stripped, so a path with spaces broke)
				$tr = "\`"$psExe\`" -NoProfile -ExecutionPolicy Bypass -File \`"$taskScriptPath\`""

				$null = schtasks /Delete /TN "$taskName" /F 2>$null
				$out = schtasks /Create /TN "$taskName" /TR $tr /SC MINUTE /MO $freqMinutes /RL HIGHEST /F /RU "SYSTEM" 2>&1

				if ($LASTEXITCODE -eq 0)
				{
					Write_Log "Scheduled Task created/updated: $taskName (every $freqMinutes minutes, SYSTEM)." "green"
				}
				else
				{
					Write_Log "Failed to create scheduled task '$taskName'. schtasks said: $out" "red"
					Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
					return
				}
			}
			catch
			{
				Write_Log "Failed to create/update Scheduled Task. Error: $_" "red"
				Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
				return
			}
		}
	}

	Write_Log "`r`n==================== Send_SMS_Key_to_Lanes Completed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Install_FUNCTIONS_Into_SMS
# ---------------------------------------------------------------------------------------------------
# Description:
#   Installs the SMS deploy artifacts into the Office folder: DEPLOY_SYS.sql (with the
#   "Multiple Functions" menu entry) and DEPLOY_MULTI_FCT.sqm (CSV / ranges like 123,234,300-305).
#   Both files are written as ANSI (Windows-1252), CRLF line endings, no BOM.
#
# Parameters:
#   - StoreNumber  [string]: kept for signature compatibility (not used in content)
#   - OfficePath   [string]: Office folder (default: $script:OfficePath, else <BasePath>\Office)
#   - Silent       [switch]: suppress Write_Log output
#
# Writes:
#   - <OfficePath>\DEPLOY_SYS.sql and <OfficePath>\DEPLOY_MULTI_FCT.sqm
# ===================================================================================================

function Install_FUNCTIONS_Into_SMS
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $false)]
		[string]$StoreNumber,
		[Parameter(Mandatory = $false)]
		[string]$OfficePath,
		[switch]$Silent
	)

	$writeLog = {
		param (
			[string]$Message,
			[string]$Color = "white"
		)
		if (-not $Silent) { Write_Log $Message $Color }
	}

	& $writeLog "`r`n==================== Starting Install_FUNCTIONS_Into_SMS ====================`r`n" "blue"

	# Registered-sign macro marker (U+00AE) built from its code point so this source stays ASCII;
	# the files are written as Windows-1252, where it is the single byte 0xAE SMS expects.
	$reg = [char]0x00AE

	# --------------------------------------------------------------------------------------------------------------
	# Resolve/validate OfficePath (prefer explicit param; fallback to common script-scoped variables)
	# --------------------------------------------------------------------------------------------------------------
	if (-not $OfficePath -or [string]::IsNullOrWhiteSpace($OfficePath))
	{
		# (the storeman ROOT is not the Office folder - before v2.7.0 this fell back to BasePath itself)
		if ($script:OfficePath) { $OfficePath = $script:OfficePath }
		elseif ($script:BasePath) { $OfficePath = Join-Path $script:BasePath 'Office' }
	}
	if (-not $OfficePath -or -not (Test-Path -LiteralPath $OfficePath))
	{
		& $writeLog "Office path not found or not provided: '$OfficePath'." "red"
		return
	}

	# --- Destination paths ---
	$DeploySysPath = Join-Path -Path $OfficePath -ChildPath "DEPLOY_SYS.sql"
	$DeployMultiFctPath = Join-Path -Path $OfficePath -ChildPath "DEPLOY_MULTI_FCT.sqm"

	# --- Encoding: ANSI (Windows-1252), no BOM ---
	$ansi = [System.Text.Encoding]::GetEncoding(1252)

	# ===============================================================================================================
	#                                             DEPLOY_SYS.sql (ALWAYS includes "Multiple Functions")
	# ===============================================================================================================
	$DeploySysContent = @"
@FMT(CMP,@dbHot(FINDFIRST,UD_DEPLOY_SYS.SQL)=,$($reg)WIZRPL(UD_RUN=0));
@FMT(CMP,@WIZGET(UD_RUN)=,'$($reg)EXEC(SQL=UD_DEPLOY_SYS)$($reg)FMT(CHR,27)');

@FMT(CMP,@TOOLS(MESSAGEDLG,"!TO KEEP THE LANE'S REFERENCE SAMPLE UP TO DATE YOU SHOULD USE THE "REFERENCE SAMPLE MECHANISM". DO YOU WANT TO CONTINUE?",,NO,YES)=1,'$($reg)FMT(CHR,27)');

@EXEC(INI=HOST_OFFICE[DEPLOY_SYS]);

@WIZRPL(STYLE=SIL);
@WIZRPL(TARGET_FILTER=@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE));

@EXEC(sqi=USERB_DEPLOY_SYS);

@WIZINIT;
@WIZMENU(ONESQM=What do you want to send,
    Functions (NEW)=DEPLOY_MULTI_FCT,
    Function link=fcz_load,
    Totalizer=tlz_load,
    Drill down pages=dril_page_load,
    Drill down files=dril_file_load,
    All system=ALL);
@WIZDISPLAY;

@WIZINIT;
@WIZTARGET(TARGET=,@FMT(CMP,"@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)=","
SELECT F1000,F1018 FROM STO_TAB WHERE F1181=1 ORDER BY F1000","
SELECT DISTINCT STO.F1000,STO.F1018 
FROM LNK_TAB LN2 JOIN LNK_TAB LNK ON LN2.F1056=LNK.F1056 AND LN2.F1057=LNK.F1057
JOIN STO_TAB STO ON STO.F1000=LNK.F1000 
WHERE STO.F1181='1' AND LN2.F1000='@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)'
ORDER BY STO.F1000"));
@WIZDISPLAY;

@FMT(CMP,@dbSelect(select distinct 1 from lnk_tab where F1000='@Wizget(Target)' and f1056='999')=,,"$($reg)EXEC(msg=!*****_can_not_deploy_system_tables_to_a_host_****);$($reg)FMT(CHR,27);")

@WIZINIT;
@WIZMENU(ACTION=Action on the target database,Add or replace=ADDRPL,Add only=ADD,Replace only=UPDATE,Clean and load=LOAD);
@WIZDISPLAY;

/* SEND ONLY ONE / MULTI / OR OTHERS */

@FMT(CMP,@wizget(ONESQM)=tlz_load,$($reg)EXEC(SQM=tlz_load));
@FMT(CMP,@wizget(ONESQM)=fcz_load,$($reg)EXEC(SQM=fcz_load));
@FMT(CMP,@wizget(ONESQM)=fct_load,$($reg)EXEC(SQM=fct_load));
@FMT(CMP,@wizget(ONESQM)=dril_file_load,$($reg)EXEC(SQM=DRIL_FILE_LOAD));
@FMT(CMP,@wizget(ONESQM)=dril_page_load,$($reg)EXEC(SQM=DRIL_PAGE_LOAD));
@FMT(CMP,@wizget(ONESQM)=DEPLOY_MULTI_FCT,$($reg)EXEC(SQM=DEPLOY_MULTI_FCT));

@FMT(CMP,@WIZGET(ONESQM)=ALL,,'$($reg)EXEC(SQM=exe_activate_accept_sys)$($reg)fmt(chr,27)');

@FMT(CMP,@wizget(tlz_load)=0,,$($reg)EXEC(SQM=tlz_load));
@FMT(CMP,@wizget(fcz_load)=0,,$($reg)EXEC(SQM=fcz_load));
@FMT(CMP,@wizget(fct_load)=0,,$($reg)EXEC(SQM=fct_load));
@FMT(CMP,@wizget(DRIL_FILE_LOAD)=0,,$($reg)EXEC(SQM=DRIL_FILE_LOAD));
@FMT(CMP,@wizget(DRIL_PAGE_LOAD)=0,,$($reg)EXEC(SQM=DRIL_PAGE_LOAD));
@FMT(CMP,@wizget(DEPLOY_MULTI_FCT)=0,,$($reg)EXEC(SQM=DEPLOY_MULTI_FCT));

@FMT(CMP,@wizget(exe_activate_accept_all)=0,,$($reg)EXEC(SQM=exe_activate_accept_sys));
@FMT(CMP,@wizget(exe_refresh_menu)=1,$($reg)EXEC(SQM=exe_refresh_menu));

@EXEC(sqi=USERE_DEPLOY_SYS);
"@

	# ===============================================================================================================
	#                                      DEPLOY_MULTI_FCT.sqm (template)
	#   Supports CSV and ranges (e.g., 123,234,300-305).
	# ===============================================================================================================
	$DeployMultiFctContent = @"
INSERT INTO HEADER_DCT VALUES
('HC','00000001','001901','001001',,,1997001,0000,1997001,0001,,'LOAD','CREATE DCT',,,,,,'1/1.0','V1.0',,);

CREATE TABLE FCT_DCT(@MAP_FROM_QUERY);

INSERT INTO HEADER_DCT VALUES
('HM','00000001','001901','001001',,,1997001,0000,1997001,0001,,'@WIZGET(ACTION)','@WIZGET(ACTION) SELECTED FUNCTIONS',,,,,,'1/1.0','V1.0','F1063',);

CREATE VIEW FCT_CHG AS SELECT @FIELDS_FROM_QUERY FROM FCT_DCT;

INSERT INTO FCT_CHG VALUES

/* EXTRACT SECTION */

@DBHOT(HOT_WIZ,PARAMTOLINE,PARAMSAV_DEPLOY_MULTI_FCT);
@FMT(CMP,'@WIZGET(TARGET)<>','$($reg)WIZRPL(TARGET_FILTER=@WIZGET(TARGET))');

@WIZINIT;
@WIZEDIT(FCT_LIST=Type ''ALL'' to deploy all functions,'Function IDs/ranges (122,129;122-129)');
@WIZDISPLAY;

@WIZINIT;
@WIZTARGET(TARGET_FILTER=,@FMT(CMP,"@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)=","
SELECT F1000,F1018 FROM STO_TAB WHERE F1181=1","
SELECT DISTINCT STO.F1000,STO.F1018
FROM LNK_TAB LN2 JOIN LNK_TAB LNK ON LN2.F1056=LNK.F1056 AND LN2.F1057=LNK.F1057
JOIN STO_TAB STO ON STO.F1000=LNK.F1000
WHERE STO.F1181='1' AND LN2.F1000='@DbHot(INI,APPLICATION.INI,DEPLOY_TARGET,HOST_OFFICE)'"));
@WIZDISPLAY;

@WIZINIT;
@WIZMENU(ACTION=Action on the target database,Add or replace=ADDRPL,Add only=ADD,Replace only=UPDATE,Create and load=LOAD);
@WIZDISPLAY;

@WIZSET(STYLE=SIL);
@WIZCLR(TARGET);
@WIZSET(FORCE_F1000=@F1056);

@dbEXEC(SET ANSI_NULLS ON)
@dbEXEC(SET QUOTED_IDENTIFIER ON)
@dbEXEC(SET ANSI_PADDING ON)
@dbEXEC(SET ANSI_WARNINGS ON)
@dbEXEC(SET CONCAT_NULL_YIELDS_NULL ON)
@dbEXEC(SET NUMERIC_ROUNDABORT OFF)
@dbEXEC(SET ARITHABORT ON)

@MAP_DEPLOY
SELECT FCT.F1056,FCT.F1056+FCT.F1057 AS F1000,@dbFld(FCT_TAB,FCT.,F1000) FROM
    (SELECT LNI.F1056,LNI.F1057,FCT.*,ROW_NUMBER() OVER (PARTITION BY FCT.F1063,LNI.F1056,LNI.F1057 ORDER BY CASE WHEN FCT.F1000='PAL' THEN 1 ELSE 2 END DESC) AS F1301
    FROM FCT_TAB FCT
    JOIN LNK_TAB LNI ON FCT.F1000=LNI.F1000
    JOIN LNK_TAB LNO ON LNI.F1056=LNO.F1056 AND LNI.F1057=LNO.F1057
    WHERE LNO.F1000='@WIZGET(TARGET_FILTER)' AND (
          UPPER(LTRIM(RTRIM('@WIZGET(FCT_LIST)'))) = 'ALL'
          OR EXISTS
          (
            SELECT 1
            FROM
            (
              SELECT
                CASE WHEN CHARINDEX('-',token)=0 AND token IS NOT NULL AND PATINDEX('%[^0-9]%',token)=0 THEN TRY_CAST(token AS INT) ELSE NULL END AS SingleNum,
                CASE WHEN CHARINDEX('-',token)>0 AND PATINDEX('%[^0-9]%',LEFT(token,CHARINDEX('-',token)-1))=0 AND PATINDEX('%[^0-9]%',SUBSTRING(token,CHARINDEX('-',token)+1,8000))=0 THEN TRY_CAST(LEFT(token,CHARINDEX('-',token)-1) AS INT) ELSE NULL END AS StartNum,
                CASE WHEN CHARINDEX('-',token)>0 AND PATINDEX('%[^0-9]%',LEFT(token,CHARINDEX('-',token)-1))=0 AND PATINDEX('%[^0-9]%',SUBSTRING(token,CHARINDEX('-',token)+1,8000))=0 THEN TRY_CAST(SUBSTRING(token,CHARINDEX('-',token)+1,8000) AS INT) ELSE NULL END AS EndNum
              FROM
              (
                SELECT T.N.value('.','nvarchar(100)') AS token
                FROM
                (
                  SELECT CAST('<i>'+REPLACE(REPLACE(REPLACE(REPLACE('@WIZGET(FCT_LIST)',CHAR(13),''),CHAR(10),''),' ',''),',','</i><i>')+'</i>' AS XML) AS xdoc
                ) X
                CROSS APPLY xdoc.nodes('/i') AS T(N)
              ) S
              WHERE ISNULL(token,'')<>''
            ) FF
            WHERE
                   (FF.SingleNum IS NOT NULL AND FCT.F1063 IS NOT NULL AND ISNUMERIC(CONVERT(VARCHAR(50),FCT.F1063))=1 AND CAST(CONVERT(VARCHAR(50),FCT.F1063) AS INT)=FF.SingleNum)
                OR (FF.StartNum IS NOT NULL AND FF.EndNum IS NOT NULL AND FF.EndNum>=FF.StartNum AND FCT.F1063 IS NOT NULL AND ISNUMERIC(CONVERT(VARCHAR(50),FCT.F1063))=1 AND CAST(CONVERT(VARCHAR(50),FCT.F1063) AS INT) BETWEEN FF.StartNum AND FF.EndNum)
          )
        )
    ) FCT
WHERE FCT.F1301=1
ORDER BY F1000,F1063;

/* RESTORE INITITAL PARAMETER POOL */
@WIZRESET;
@DBHOT(HOT_WIZ,LINETOPARAM,PARAMSAV_DEPLOY_MULTI_FCT);
@DBHOT(HOT_WIZ,CLR,PARAMSAV_DEPLOY_MULTI_FCT);
"@

	# ===============================================================================================================
	#                                                  WRITE FILES
	#   clear read-only/hidden -> normalize to CRLF -> write ANSI (no BOM) when changed -> log
	# ===============================================================================================================

	# Clears read-only / hidden BEFORE writing (a read-only file made the write fail), and leaves a
	# file that already has this exact content untouched (no needless rewrite at every launch)
	$writeAnsi = {
		param ([string]$Path, [string]$Content)
		$norm = [regex]::Replace($Content, "(`r)?`n", "`r`n")
		if ([System.IO.File]::Exists($Path))
		{
			[System.IO.File]::SetAttributes($Path, [System.IO.FileAttributes]::Normal)
			if ([System.IO.File]::ReadAllText($Path, $ansi) -ceq $norm) { return $false }
		}
		[System.IO.File]::WriteAllText($Path, $norm, $ansi)
		return $true
	}

	foreach ($out in @(
			@($DeploySysPath, $DeploySysContent, 'DEPLOY_SYS.sql'),
			@($DeployMultiFctPath, $DeployMultiFctContent, 'DEPLOY_MULTI_FCT.sqm (All + CSV + ranges)')))
	{
		try
		{
			if (& $writeAnsi $out[0] $out[1]) { & $writeLog "Wrote '$($out[2])' to '$OfficePath'." "green" }
			else { & $writeLog "'$($out[2])' in '$OfficePath' is already current." "gray" }
		}
		catch { & $writeLog "Failed to write '$($out[2])'. Error: $_" "red" }
	}
	& $writeLog "`r`n==================== Install_FUNCTIONS_Into_SMS Completed ====================`r`n" "blue"
}

# ===================================================================================================
#                              FUNCTION: Insert_Test_Item
# ---------------------------------------------------------------------------------------------------
# Description:
#   Inserts a deterministic test item into the store database: picks one of three test PLUs
#   (the first that is unused or already a test item), deletes any existing rows for it, then
#   inserts fixed rows into SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB and SCL_TXT_TAB.
#   Uses System.Data.SqlClient (works without the SqlServer / SQLPS module) over one connection,
#   and does the deletes + inserts in one transaction: a failure leaves the database unchanged
#   instead of half a test item.
#
# Parameters:
#   - ConnectionString         [string]: defaults to the FunctionResults connection string
#   - Silent                   [switch]: suppress Write_Log output (startup job: no Write_Log there)
#   - QueryTimeoutSeconds      [int]: SQL query timeout in seconds (default 15)
#   - ConnectionTimeoutSeconds [int]: SQL connect timeout in seconds (default 5)
#
# Variables:
#   - $script:FunctionResults: reads ConnectionString, DBSERVER, DBNAME
#
# Notes:
#   - Self-contained: the startup warmup runs it in a plain Start-Job with only its own body.
#
# Writes:
#   - store database (SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB, SCL_TXT_TAB)
# ===================================================================================================

function Insert_Test_Item
{
	[CmdletBinding()]
	param (
		[string]$ConnectionString = $script:FunctionResults['ConnectionString'],
		[switch]$Silent,
		[int]$QueryTimeoutSeconds = 15,
		[int]$ConnectionTimeoutSeconds = 5
	)

	$writeLog = {
		param (
			[string]$Message,
			[string]$Color = "white"
		)
		if (-not $Silent) { Write_Log $Message $Color }
	}

	# No stored connection string: Windows login to DBSERVER / DBNAME
	if ([string]::IsNullOrWhiteSpace($ConnectionString) -or $ConnectionString -eq 'N/A')
	{
		$dbServer = [string]$script:FunctionResults['DBSERVER']
		$dbName = [string]$script:FunctionResults['DBNAME']
		$ConnectionString = $null
		if ($dbServer -and $dbName -and $dbServer -ne 'N/A' -and $dbName -ne 'N/A')
		{
			$ConnectionString = "Server=$dbServer;Database=$dbName;Integrated Security=True;TrustServerCertificate=True;"
		}
	}
	if (-not $ConnectionString)
	{
		& $writeLog "Insert_Test_Item: no connection string and no DBSERVER / DBNAME. Run Get_Store_And_Database_Info first." "red"
		return
	}
	try
	{
		$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder $ConnectionString
		$csb.PSBase.ConnectTimeout = $ConnectionTimeoutSeconds
		$ConnectionString = $csb.ConnectionString
	}
	catch
	{
		if ($ConnectionString -notmatch '(?i)\bConnect\s*Timeout\s*=') { $ConnectionString = $ConnectionString.Trim().TrimEnd(';') + ";Connect Timeout=$ConnectionTimeoutSeconds" }
	}

	& $writeLog "`r`n==================== Starting Insert_Test_Item ====================`r`n" "blue"

	$conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
	try { $conn.Open() }
	catch
	{
		& $writeLog "Insert_Test_Item: cannot connect to the store database: $($_.Exception.Message)" "red"
		& $writeLog "`r`n==================== Insert_Test_Item Aborted (No Connectivity) ====================`r`n" "blue"
		$conn.Dispose()
		return
	}

	try
	{
		$tx = $null
		$runSql = {
			param ([string]$Sql, [switch]$Scalar)
			$cmd = $conn.CreateCommand()
			$cmd.CommandText = $Sql
			$cmd.CommandTimeout = $QueryTimeoutSeconds
			if ($tx) { $cmd.Transaction = $tx }
			try
			{
				if ($Scalar)
				{
					$v = $cmd.ExecuteScalar()
					if ($v -is [DBNull]) { return $null }
					return $v
				}
				[void]$cmd.ExecuteNonQuery()
			}
			finally { $cmd.Dispose() }
		}

		# ------------------------------------------------------------------------------------------
		# Test PLU: the first candidate that is unused or already a test item (description
		# contains "test" / "tecnica"). A candidate whose lookup FAILS is skipped - never assume
		# a PLU is free when it could not be checked (it may be a real item).
		# ------------------------------------------------------------------------------------------
		$PLU = $null
		$TestF267 = $null
		foreach ($candidate in @(
				@{ Plu = '0020077700000'; F267 = 777; Kind = 'preferred' },
				@{ Plu = '0020777700000'; F267 = 7777; Kind = 'alternative' },
				@{ Plu = '0027777700000'; F267 = 77777; Kind = 'fallback' }))
		{
			try
			{
				$posDesc = [string](& $runSql "SELECT TOP 1 F02 FROM POS_TAB WHERE F01 = '$($candidate.Plu)'" -Scalar)
				$objDesc = [string](& $runSql "SELECT TOP 1 F29 FROM OBJ_TAB WHERE F01 = '$($candidate.Plu)'" -Scalar)
				$usable = ($posDesc -match '(?i)test|tecnica') -or ($objDesc -match '(?i)test|tecnica') -or
				([string]::IsNullOrWhiteSpace($posDesc) -and [string]::IsNullOrWhiteSpace($objDesc))
			}
			catch
			{
				& $writeLog ("Could not check PLU {0}: {1} - skipping it." -f $candidate.Plu, $_.Exception.Message) "yellow"
				$usable = $false
			}
			if ($usable)
			{
				$PLU = $candidate.Plu
				$TestF267 = $candidate.F267
				& $writeLog ("Using {0} PLU: {1} with F267: {2}" -f $candidate.Kind, $PLU, $TestF267) "green"
				break
			}
		}
		if (-not $PLU)
		{
			& $writeLog "`r`n==================== Insert_Test_Item Completed (No Data) ====================`r`n" "blue"
			return
		}

		$now = Get-Date
		$nowFull = $now.ToString("yyyy-MM-dd HH:mm:ss.fff")
		$nowDate = $now.ToString("yyyy-MM-dd 00:00:00.000")
		$F29 = 'Tecnica Test Item'

		# Old rows first, so the inserts are deterministic. The scale text row is only removed
		# when no other item uses that F267 (it could be a real one).
		$steps = @(
			@('delete SCL_TAB', "DELETE FROM SCL_TAB     WHERE F01 = '$PLU'"),
			@('delete OBJ_TAB', "DELETE FROM OBJ_TAB     WHERE F01 = '$PLU'"),
			@('delete POS_TAB', "DELETE FROM POS_TAB     WHERE F01 = '$PLU'"),
			@('delete PRICE_TAB', "DELETE FROM PRICE_TAB   WHERE F01 = '$PLU'"),
			@('delete SCL_TXT_TAB', "DELETE FROM SCL_TXT_TAB WHERE F267 = $TestF267 AND F267 NOT IN (SELECT F267 FROM SCL_TAB WHERE F01 <> '$PLU' AND F267 IS NOT NULL)"),
			@('SCL_TAB', "INSERT INTO SCL_TAB (F01, F1000, F902, F1001, F253, F258, F264, F267, F1952, F1964, F2581, F2582) VALUES ('$PLU', 'PAL', 'MANUAL', 1, '$nowFull', 10, 7, $TestF267, 'Test Descriptor 2', '001', 'Test Descriptor 3', 'Test Descriptor 4')"),
			@('OBJ_TAB', "INSERT INTO OBJ_TAB (F01, F902, F1001, F21, F29, F270, F1118, F1959) VALUES ('$PLU', '00001153', 0, 1, '$F29', 123.45, '001', '001')"),
			@('POS_TAB', "INSERT INTO POS_TAB (F01, F1000, F902, F1001, F02, F09, F79, F80, F82, F104, F115, F176, F178, F217, F1964, F2119) VALUES ('$PLU', 'PAL', 'MANUAL', 0, 'Tecnica Test Item', '$nowDate', '1', '1', '1', '0', '0', '1', '1', 1.0, '001', '1')"),
			@('PRICE_TAB', "INSERT INTO PRICE_TAB (F01, F1000, F126, F902, F1001, F21, F30, F31, F113, F1006, F1007, F1008, F1009, F1803) VALUES ('$PLU', 'PAL', 1, 'MANUAL', 0, 1, 777.77, 1, 'REG', 1, 777.77, '$nowDate', '1858', 1.0)"),
			@('SCL_TXT_TAB', "INSERT INTO SCL_TXT_TAB (F267, F1000, F253, F297, F902, F1001, F1836) VALUES ($TestF267, 'PAL', '$nowFull', 'Ingredients Test', 'MANUAL', 0, 'Tecnica Test Item')")
		)

		$tx = $conn.BeginTransaction()
		$failed = $null
		foreach ($s in $steps)
		{
			try { & $runSql $s[1] }
			catch
			{
				$failed = "{0}: {1}" -f $s[0], $_.Exception.Message
				break
			}
		}
		if ($failed)
		{
			try { $tx.Rollback() }
			catch { }
			& $writeLog "Test item NOT inserted - $failed (rolled back, the database is unchanged)." "red"
			& $writeLog "`r`n==================== Insert_Test_Item Aborted ====================`r`n" "blue"
			return
		}
		$tx.Commit()
		& $writeLog "Test item $PLU inserted into SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB and SCL_TXT_TAB." "green"
		& $writeLog "`r`n==================== Insert_Test_Item Completed ====================`r`n" "blue"
	}
	catch
	{
		& $writeLog "Insert_Test_Item failed: $($_.Exception.Message)" "red"
	}
	finally
	{
		$conn.Dispose()
	}
}

# ===================================================================================================
#                              FUNCTION: Troubleshoot_ScaleCommApp
# ---------------------------------------------------------------------------------------------------
# Description:
#   Audits and optionally repairs the ScaleCommApp *.exe.config files: sets StoreName (the
#   store name resolved by Get_Store_And_Database_Info), FistScaleIP (first active scale IP),
#   SQL_InstanceName, SQL_databaseName and Recs_BatchSendFull; restores the deploy line in
#   Office\DEPLOY_CHG.sql and removes the "Update_Scales_Specials_Task_Minutes" task if present
#   (both through Fix_Deploy_CHG).
#
# Parameters:
#   - ConfigPath          [string]: specific config file to target
#   - ScaleCommRoots      [string[]]: candidate ScaleCommApp install roots (default $script:ScaleCommRoots)
#   - AllMatches          [switch]: fix every matching config (default: first only)
#   - NoAutoFix           [switch]: audit only, do not save
#   - Force               [switch]: overwrite without confirmation
#   - ExpectedStoreName   [string]: expected StoreName (default: $script:FunctionResults['StoreName'])
#   - ExpectedSqlInstance / ExpectedDatabase [string]: expected SQL instance (DBSERVER) / database (DBNAME)
#   - OfficePath          [string]: Office folder for DEPLOY_CHG.sql (default $script:OfficePath)
#
# Variables:
#   - $script:FunctionResults: DBNAME, DBSERVER, ConnectionString, StoreName; $script:ScaleCommRoots
#
# Writes:
#   - ScaleCommApp *.exe.config files (UTF-8 no BOM, .bak backups) and Office\DEPLOY_CHG.sql
# ===================================================================================================

function Troubleshoot_ScaleCommApp
{
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
		[string]$ConfigPath,
		[string[]]$ScaleCommRoots = $script:ScaleCommRoots,
		[switch]$AllMatches,
		[switch]$NoAutoFix,
		[switch]$Force,
		[string]$ExpectedStoreName,
		[string]$ExpectedSqlInstance,
		[string]$ExpectedDatabase,
		# NEW: where Office\DEPLOY_CHG.sql lives. If not passed, uses $script:OfficePath when available.
		[string]$OfficePath
	)

	Write_Log "`r`n==================== Starting Troubleshoot_ScaleCommApp ====================`r`n" "blue"

	# ------------------------------------------------------------------------------------------------
	# Resolve expectations & connection (same behavior; no ternary)
	# ------------------------------------------------------------------------------------------------
	$dbName = $script:FunctionResults['DBNAME']
	$server = $script:FunctionResults['DBSERVER']
	$connStr = $script:FunctionResults['ConnectionString']

	if (-not $ExpectedDatabase -and $dbName) { $ExpectedDatabase = $dbName }
	if (-not $ExpectedSqlInstance -and $server) { $ExpectedSqlInstance = $server }
	# Store name: resolved once by Get_Store_And_Database_Info (system.ini first, INFO WIN.INI
	# fallback) - no second parse here.

	if (-not $connStr -and $server -and $dbName)
	{
		# Build connection string inline (no helpers)
		$sqlUser = $script:FunctionResults['SQL_UserID']
		$sqlPwd = $script:FunctionResults['SQL_UserPwd']
		if ($sqlUser -and $sqlPwd)
		{
			$connStr = "Server=$server;Database=$dbName;User ID=$sqlUser;Password=$sqlPwd;TrustServerCertificate=True;"
		}
		else
		{
			$connStr = "Server=$server;Database=$dbName;Integrated Security=SSPI;TrustServerCertificate=True;"
		}
		$script:FunctionResults['ConnectionString'] = $connStr
	}

	# ------------------------------------------------------------------------------------------------
	# Active NIC prefix (fallback) - logged ONCE (inline)
	# ------------------------------------------------------------------------------------------------
	$ActiveNicPrefix = $null
	try
	{
		# Collect server IPv4s (for subnet check)
		$serverIPv4s = @()
		try
		{
			[System.Net.Dns]::GetHostAddresses((($server -replace '^(?i)tcp:', '') -split '[\\,]')[0]) | ForEach-Object {
				if ($_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $serverIPv4s += $_.ToString() }
			}
		}
		catch { }

		# Get all UP NICs
		$allNics = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
		Where-Object { $_.OperationalStatus -eq [System.Net.NetworkInformation.OperationalStatus]::Up }

		# Try to pick an interface on the same subnet as SQL server (inline mask compare)
		$pickedIP = $null
		foreach ($ni in $allNics)
		{
			$ipProps = $ni.GetIPProperties()
			foreach ($ua in $ipProps.UnicastAddresses)
			{
				if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
				if (-not $ua.IPv4Mask) { continue }

				foreach ($srv in $serverIPv4s)
				{
					$ok = $false
					try
					{
						$a = [System.Net.IPAddress]::Parse($ua.Address.ToString()).GetAddressBytes()
						$b = [System.Net.IPAddress]::Parse($srv).GetAddressBytes()
						$m = [System.Net.IPAddress]::Parse($ua.IPv4Mask.ToString()).GetAddressBytes()
						$ok = $true
						for ($i = 0; $i -lt 4; $i++) { if (($a[$i] -band $m[$i]) -ne ($b[$i] -band $m[$i])) { $ok = $false; break } }
					}
					catch { $ok = $false }
					if ($ok) { $pickedIP = $ua.Address.ToString(); break }
				}
				if ($pickedIP) { break }
			}
			if ($pickedIP) { break }
		}

		# Gateway-backed interface fallback
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				$hasGw = $false
				foreach ($gw in $ipProps.GatewayAddresses)
				{
					if ($gw.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { $hasGw = $true; break }
				}
				if (-not $hasGw) { continue }
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$pickedIP = $ua.Address.ToString(); break
				}
				if ($pickedIP) { break }
			}
		}

		# RFC1918 preference fallback
		if (-not $pickedIP)
		{
			foreach ($ni in $allNics)
			{
				$ipProps = $ni.GetIPProperties()
				foreach ($ua in $ipProps.UnicastAddresses)
				{
					if ($ua.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { continue }
					$ip = $ua.Address.ToString()
					if ($ip.StartsWith('10.') -or $ip.StartsWith('192.168.')) { $pickedIP = $ip; break }
					if ($ip.StartsWith('172.'))
					{
						$parts = $ip.Split('.')
						if ($parts.Length -ge 2)
						{
							$sec = 0
							if ([int]::TryParse($parts[1], [ref]$sec)) { if ($sec -ge 16 -and $sec -le 31) { $pickedIP = $ip; break } }
						}
					}
				}
				if ($pickedIP) { break }
			}
		}

		# Derive /24 prefix
		if ($pickedIP) { $p = $pickedIP.Split('.'); if ($p.Length -ge 3) { $ActiveNicPrefix = "$($p[0]).$($p[1]).$($p[2])" } }
	}
	catch { }

	if ($ActiveNicPrefix) { Write_Log "Active NIC /24 prefix: $ActiveNicPrefix" "cyan" }
	else { Write_Log "Active NIC /24 prefix not detected (fallback may be unavailable)." "yellow" }

	# ------------------------------------------------------------------------------------------------
	# FIRST active scale full IP (robust) - logged ONCE (inline)
	# ------------------------------------------------------------------------------------------------
	$firstScaleIP = $null
	if ($server -and $dbName -and $connStr)
	{

		# Query
		$q = @"
SELECT IPNetwork, IPDevice, ScaleCode, Active
FROM dbo.TBS_SCL_ver520
ORDER BY
    CASE WHEN TRY_CAST(ScaleCode AS INT) IS NULL THEN 1 ELSE 0 END,
    TRY_CAST(ScaleCode AS INT),
    ScaleCode;
"@

		# Execute query (inline; no helpers)
		$rows = $null
		try
		{
			# Built-in SqlClient (no SQL module import for one query)
			$conn = New-Object System.Data.SqlClient.SqlConnection $connStr
			try
			{
				$conn.Open()
				$cmd = $conn.CreateCommand(); $cmd.CommandText = $q; $cmd.CommandTimeout = 30
				$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
				$dt = New-Object System.Data.DataTable
				[void]$da.Fill($dt)
				$rows = $dt
			}
			finally { $conn.Dispose() }
		}
		catch { $rows = $null }

		# Process rows
		$hasRows = $false
		if ($rows -is [System.Data.DataTable]) { if ($rows.Rows.Count -gt 0) { $hasRows = $true } }
		elseif ($rows) { if ($rows.Count -gt 0) { $hasRows = $true } }

		if ($hasRows)
		{
			Write_Log "Loaded rows from TBS_SCL_ver520 for first-scale detection." "cyan"

			# Normalize enumeration
			$enum = @()
			if ($rows -is [System.Data.DataTable]) { $enum = $rows.Rows }
			else { $enum = $rows }

			# Prefix frequency map (inline)
			$pfxCounts = @{ }
			foreach ($r in $enum)
			{
				$ipn = $null
				if ($null -ne $r.IPNetwork) { $ipn = [string]$r.IPNetwork }
				if ($ipn)
				{
					$ipn = $ipn.Trim()
					$pref = $null
					$m = [regex]::Match($ipn, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if ($m.Success) { $pref = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)" }
					else
					{
						$m4 = [regex]::Match($ipn, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\s*$')
						if ($m4.Success) { $pref = "$($m4.Groups[1].Value).$($m4.Groups[2].Value).$($m4.Groups[3].Value)" }
					}
					if ($pref)
					{
						if (-not $pfxCounts.ContainsKey($pref)) { $pfxCounts[$pref] = 0 }
						$pfxCounts[$pref] = $pfxCounts[$pref] + 1
					}
				}
			}

			$MostCommonPrefix = $null
			if ($pfxCounts.Count -gt 0)
			{
				$MostCommonPrefix = ($pfxCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
				Write_Log "Most common IPNetwork prefix in table: $MostCommonPrefix" "cyan"
			}

			# Filter "active" rows (inline)
			$candidates = @()
			foreach ($r in $enum)
			{
				$act = $null
				if ($null -ne $r.Active) { $act = ([string]$r.Active).Trim().ToUpper() }
				else { $act = '' }
				$isActive = $false
				if ($act -eq '') { $isActive = $true }
				elseif ($act -eq 'Y' -or $act -eq '1' -or $act -eq 'TRUE') { $isActive = $true }
				if ($isActive) { $candidates += $r }
			}

			# Try 1: exact per-row prefix + device
			foreach ($r in $candidates)
			{
				$pref = $null
				$ipn = $null
				if ($null -ne $r.IPNetwork) { $ipn = ([string]$r.IPNetwork).Trim() }
				if ($ipn)
				{
					$m = [regex]::Match($ipn, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})(?:\.(?:\d{1,3}|\*|x|X)?)?\s*$')
					if ($m.Success) { $pref = "$($m.Groups[1].Value).$($m.Groups[2].Value).$($m.Groups[3].Value)" }
					if (-not $pref)
					{
						$m4 = [regex]::Match($ipn, '^\s*(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\s*$')
						if ($m4.Success)
						{
							$pref = "$($m4.Groups[1].Value).$($m4.Groups[2].Value).$($m4.Groups[3].Value)"
							if (-not $r.IPDevice) { $r.IPDevice = $m4.Groups[4].Value }
						}
					}
				}

				$devOct = $null
				if ($null -ne $r.IPDevice)
				{
					$tmp = 0
					if ([int]::TryParse(([string]$r.IPDevice).Trim(), [ref]$tmp))
					{
						if ($tmp -ge 0 -and $tmp -le 255) { $devOct = $tmp }
					}
				}

				if ($pref -and ($null -ne $devOct)) { $firstScaleIP = "$pref.$devOct"; break }
			}

			# Try 2: table-most-common prefix + device
			if (-not $firstScaleIP -and $MostCommonPrefix)
			{
				foreach ($r in $candidates)
				{
					$devOct = $null
					if ($null -ne $r.IPDevice)
					{
						$tmp = 0
						if ([int]::TryParse(([string]$r.IPDevice).Trim(), [ref]$tmp))
						{
							if ($tmp -ge 0 -and $tmp -le 255) { $devOct = $tmp }
						}
					}
					if ($null -ne $devOct) { $firstScaleIP = "$MostCommonPrefix.$devOct"; break }
				}
			}

			# Try 3: active NIC /24 prefix + device
			if (-not $firstScaleIP -and $ActiveNicPrefix)
			{
				foreach ($r in $candidates)
				{
					$devOct = $null
					if ($null -ne $r.IPDevice)
					{
						$tmp = 0
						if ([int]::TryParse(([string]$r.IPDevice).Trim(), [ref]$tmp))
						{
							if ($tmp -ge 0 -and $tmp -le 255) { $devOct = $tmp }
						}
					}
					if ($null -ne $devOct) { $firstScaleIP = "$ActiveNicPrefix.$devOct"; break }
				}
			}

			if ($firstScaleIP) { Write_Log "First scale IP resolved: $firstScaleIP" "cyan" }
			else { Write_Log "First scale IP NOT resolved (need prefix + device)." "yellow" }
		}
		else
		{
			Write_Log "Could not read rows from TBS_SCL_ver520 (DB issue or empty table)." "yellow"
		}
	}
	else
	{
		Write_Log "DB connection info incomplete; first-scale detection skipped." "yellow"
	}

	# ------------------------------------------------------------------------------------------------
	# Locate target configs
	# ------------------------------------------------------------------------------------------------
	$targets = @()
	if ($ConfigPath)
	{
		if (Test-Path -LiteralPath $ConfigPath) { $targets += (Get-Item -LiteralPath $ConfigPath) }
		else { Write_Log "Config path not found: $ConfigPath" "red"; return }
	}
	else
	{
		foreach ($root in $ScaleCommRoots)
		{
			if (Test-Path -LiteralPath $root)
			{
				$found = Get-ChildItem -LiteralPath $root -File -Filter '*.exe.config' -ErrorAction SilentlyContinue
				if ($found) { $targets += $found }
			}
		}
	}
	$targets = $targets | Sort-Object FullName -Unique
	if (-not $targets -or $targets.Count -eq 0) { Write_Log "No *.exe.config found under ScaleCommApp roots." "red"; return }
	if (-not $AllMatches -and $targets.Count -gt 1)
	{
		Write_Log ("Multiple configs found; using first. Use -AllMatches to update all." + "`r`n - " + (($targets | ForEach-Object FullName) -join "`r`n - ")) "yellow"
		$targets = @($targets[0])
	}

	# ------------------------------------------------------------------------------------------------
	# ONE-TIME "Desired Settings" + "Existing (first file)" snapshot (inline getters)
	# ------------------------------------------------------------------------------------------------
	$previewPath = $targets[0].FullName
	$previewXml = $null
	try { [xml]$previewXml = Get-Content -LiteralPath $previewPath -Raw -ErrorAction Stop }
	catch { }

	$curStorePreview = $null; $curFirstIPPreview = $null; $curSqlInstPreview = $null; $curDbNamePreview = $null; $curBatchPreview = $null
	if ($previewXml)
	{
		$cfgPrev = $previewXml.configuration; if (-not $cfgPrev) { $cfgPrev = $previewXml.DocumentElement }
		if ($cfgPrev)
		{
			$asPrev = $cfgPrev.appSettings
			if ($asPrev)
			{
				$n = $asPrev.SelectSingleNode("add[@key='StoreName']"); if ($n) { $curStorePreview = [string]$n.GetAttribute('value') }
				$n = $asPrev.SelectSingleNode("add[@key='FistScaleIP']"); if ($n) { $curFirstIPPreview = [string]$n.GetAttribute('value') }
				$n = $asPrev.SelectSingleNode("add[@key='SQL_InstanceName']"); if ($n) { $curSqlInstPreview = [string]$n.GetAttribute('value') }
				$n = $asPrev.SelectSingleNode("add[@key='SQL_databaseName']"); if ($n) { $curDbNamePreview = [string]$n.GetAttribute('value') }
				$n = $asPrev.SelectSingleNode("add[@key='Recs_BatchSendFull']"); if ($n) { $curBatchPreview = [string]$n.GetAttribute('value') }
			}
		}
	}

	$dStore = '<unchanged/skip>'; if ($ExpectedStoreName) { $dStore = $ExpectedStoreName }
	$dIP = '<skip>'; if ($firstScaleIP) { $dIP = $firstScaleIP }
	$dInst = '<unchanged/skip>'; if ($ExpectedSqlInstance) { $dInst = $ExpectedSqlInstance }
	$dDb = '<unchanged/skip>'; if ($ExpectedDatabase) { $dDb = $ExpectedDatabase }

	Write_Log "Desired Settings:" "blue"
	Write_Log ("  StoreName          : {0}" -f $dStore) "green"
	Write_Log ("  FistScaleIP        : {0}" -f $dIP) "green"
	Write_Log ("  SQL_InstanceName   : {0}" -f $dInst) "green"
	Write_Log ("  SQL_databaseName   : {0}" -f $dDb) "green"
	Write_Log ("  Recs_BatchSendFull : 10000") "green"

	$leafPreview = Split-Path -Leaf $previewPath
	Write_Log ("Existing (first file: {0}):" -f $leafPreview) "blue"
	if (-not $curStorePreview) { $curStorePreview = '<missing>' }
	if (-not $curFirstIPPreview) { $curFirstIPPreview = '<missing>' }
	if (-not $curSqlInstPreview) { $curSqlInstPreview = '<missing>' }
	if (-not $curDbNamePreview) { $curDbNamePreview = '<missing>' }
	if (-not $curBatchPreview) { $curBatchPreview = '<missing>' }
	Write_Log ("  StoreName          : {0}" -f $curStorePreview) "yellow"
	Write_Log ("  FistScaleIP        : {0}" -f $curFirstIPPreview) "yellow"
	Write_Log ("  SQL_InstanceName   : {0}" -f $curSqlInstPreview) "yellow"
	Write_Log ("  SQL_databaseName   : {0}" -f $curDbNamePreview) "yellow"
	Write_Log ("  Recs_BatchSendFull : {0}" -f $curBatchPreview) "yellow"

	# ------------------------------------------------------------------------------------------------
	# Apply to all targets (quiet per-file; track counts + saved list)
	# ------------------------------------------------------------------------------------------------
	$autoFix = (-not $NoAutoFix.IsPresent)
	[int]$audited = 0
	[int]$updated = 0
	[int]$skipped = 0
	$savedFiles = New-Object System.Collections.Generic.List[string]

	foreach ($t in $targets)
	{
		$audited++
		$path = $t.FullName

		# Load XML
		$xml = $null
		try { [xml]$xml = Get-Content -LiteralPath $path -Raw -ErrorAction Stop }
		catch { $skipped++; continue }
		$cfg = $xml.configuration; if (-not $cfg) { $cfg = $xml.DocumentElement }
		if (-not $cfg) { $skipped++; continue }
		$as = $cfg.appSettings
		if (-not $as) { $skipped++; continue }

		# READ current values (inline)
		$curStore = $null; $curFirstIP = $null; $curInst = $null; $curDb = $null; $curBatch = $null
		$n = $as.SelectSingleNode("add[@key='StoreName']"); if ($n) { $curStore = [string]$n.GetAttribute('value') }
		$n = $as.SelectSingleNode("add[@key='FistScaleIP']"); if ($n) { $curFirstIP = [string]$n.GetAttribute('value') }
		$n = $as.SelectSingleNode("add[@key='SQL_InstanceName']"); if ($n) { $curInst = [string]$n.GetAttribute('value') }
		$n = $as.SelectSingleNode("add[@key='SQL_databaseName']"); if ($n) { $curDb = [string]$n.GetAttribute('value') }
		$n = $as.SelectSingleNode("add[@key='Recs_BatchSendFull']"); if ($n) { $curBatch = [string]$n.GetAttribute('value') }

		$changedHere = $false

		# SET (inline create-if-missing)
		if ($ExpectedStoreName -and $curStore -ne $ExpectedStoreName)
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set StoreName'))
			{
				$n = $as.SelectSingleNode("add[@key='StoreName']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'StoreName'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', [string]$ExpectedStoreName)
				$changedHere = $true
			}
		}

		if ($firstScaleIP -and $curFirstIP -ne $firstScaleIP)
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set FistScaleIP'))
			{
				$n = $as.SelectSingleNode("add[@key='FistScaleIP']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'FistScaleIP'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', [string]$firstScaleIP)
				$changedHere = $true
			}
		}

		if ($ExpectedSqlInstance -and $curInst -ne $ExpectedSqlInstance)
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set SQL_InstanceName'))
			{
				$n = $as.SelectSingleNode("add[@key='SQL_InstanceName']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'SQL_InstanceName'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', [string]$ExpectedSqlInstance)
				$changedHere = $true
			}
		}

		if ($ExpectedDatabase -and $curDb -ne $ExpectedDatabase)
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set SQL_databaseName'))
			{
				$n = $as.SelectSingleNode("add[@key='SQL_databaseName']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'SQL_databaseName'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', [string]$ExpectedDatabase)
				$changedHere = $true
			}
		}

		if ($curBatch -ne '10000')
		{
			if ($autoFix -and $PSCmdlet.ShouldProcess($path, 'Set Recs_BatchSendFull=10000'))
			{
				$n = $as.SelectSingleNode("add[@key='Recs_BatchSendFull']")
				if (-not $n) { $n = $xml.CreateElement('add'); $null = $n.SetAttribute('key', 'Recs_BatchSendFull'); $null = $as.AppendChild($n) }
				$null = $n.SetAttribute('value', '10000')
				$changedHere = $true
			}
		}

		# Save if changed
		if ($changedHere -and $autoFix)
		{
			try
			{
				$item = Get-Item -LiteralPath $path -ErrorAction Stop
				if ($item.Attributes -band [IO.FileAttributes]::ReadOnly)
				{
					if ($Force) { $item.Attributes = ($item.Attributes -bxor [IO.FileAttributes]::ReadOnly) }
					else { $skipped++; continue }
				}
			}
			catch { $skipped++; continue }

			try
			{
				$dir = Split-Path -LiteralPath $path -Parent
				$leaf = Split-Path -LiteralPath $path -Leaf
				$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
				$bak = Join-Path $dir "$leaf.$stamp.bak"
				Copy-Item -LiteralPath $path -Destination $bak -Force
			}
			catch { }

			try
			{
				$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
				$sw = New-Object System.IO.StreamWriter($path, $false, $utf8NoBom)
				try { $xml.Save($sw) }
				finally { $sw.Close() }
				$updated++; $savedFiles.Add((Split-Path -Leaf $path)) | Out-Null
			}
			catch { $skipped++ }
		}
	}

	# ==================================================================================================
	# NEW: Restore DEPLOY_CHG.sql deploy line (integrated Fix_Deploy_CHG)
	# ==================================================================================================
	if (-not $OfficePath -and $script:OfficePath) { $OfficePath = $script:OfficePath }
	if (-not $OfficePath)
	{
		Write_Log "Office path not provided and \$script:OfficePath not available; skipping DEPLOY_CHG.sql restore." "yellow"
	}
	elseif ($PSCmdlet.ShouldProcess((Join-Path $OfficePath 'DEPLOY_CHG.sql'), 'Restore DEPLOY_CHG @EXEC line'))
	{
		[void](Fix_Deploy_CHG -OfficePath $OfficePath -Quiet)
	}
	# ------------------------------------------------------------------------------------------------
	# Final summary
	# ------------------------------------------------------------------------------------------------
	$sumIP = 'n/a'
	if ($firstScaleIP) { $sumIP = $firstScaleIP }

	Write_Log ("`r`nTroubleshoot_ScaleCommApp: Audited={0}, Updated={1}, Skipped={2}, FirstScaleIP={3}" -f $audited, $updated, $skipped, $sumIP) "blue"
	if ($updated -gt 0) { Write_Log ("Saved files: {0}" -f ($savedFiles -join ', ')) "green" }

	Write_Log "`r`n==================== Troubleshoot_ScaleCommApp Completed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Get_All_VNC_Passwords
# ---------------------------------------------------------------------------------------------------
# Description:
#   Scans all lanes, backoffices and scales for the UltraVNC configuration file (UltraVNC.ini,
#   any case) in the standard install locations, extracts the encrypted VNC password and
#   returns a machine-name -> password hashtable. Uses a fixed default password for nodes
#   where no UltraVNC.ini can be read. Online lanes / backoffices / scales are read in
#   parallel in a runspace pool (20 threads, 2 minutes in total); Bizerba scales one by one
#   (they need the bizuser credential).
#
# Parameters:
#   - LaneNumToMachineName       [hashtable]: lane number -> machine name
#   - ScaleCodeToIPInfo          [hashtable]: scale code -> scale info object
#   - BackofficeNumToMachineName [hashtable]: backoffice number -> machine name
#
# Variables:
#   - $script:DefaultVNCPassword: fixed default for stock-password nodes (e.g. Ishida)
#
# Returns:
#   Hashtable MachineName => encrypted password (or $null when not found).
#
# Writes:
#   - $script:FunctionResults['AllVNCPasswords']
# ===================================================================================================

function Get_All_VNC_Passwords
{
	param (
		[Parameter(Mandatory = $false)]
		[hashtable]$LaneNumToMachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$ScaleCodeToIPInfo,
		[Parameter(Mandatory = $false)]
		[hashtable]$BackofficeNumToMachineName
	)

	# Default VNC password for lanes, backoffices, and scales with fixed passwords (e.g., Ishida)
	$DefaultVNCPassword = "4330df922eb03b6e"
	$script:DefaultVNCPassword = $DefaultVNCPassword

	$uvncFolders = @(
		"C$\Program Files\uvnc bvba\UltraVNC",
		"C$\Program Files (x86)\uvnc bvba\UltraVNC"
	)
	$AllVNCPasswords = @{ }

	# 1. Build main node list and tag brands for scales
	$NodeList = @()
	$BizerbaScales = @()
	if ($LaneNumToMachineName) { $NodeList += $LaneNumToMachineName.Values | Where-Object { $_ } }
	if ($ScaleCodeToIPInfo)
	{
		$uniqueScaleObjs = @{ }
		foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
		{
			$scaleObj = $kv.Value
			# Compute a unique identifier, e.g., IP or some combination
			$ip = $null
			if ($scaleObj.FullIP) { $ip = $scaleObj.FullIP }
			elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice) { $ip = (([string]$scaleObj.IPNetwork).TrimEnd('.') + '.' + [string]$scaleObj.IPDevice) }
			# Add only once!
			if ($ip -and -not $uniqueScaleObjs.ContainsKey($ip))
			{
				$uniqueScaleObjs[$ip] = $scaleObj
			}
		}
		foreach ($ip in $uniqueScaleObjs.Keys)
		{
			$scaleObj = $uniqueScaleObjs[$ip]
			$ip = $null
			$isIshida = $false
			$isBizerba = $false
			if ($scaleObj -is [string]) { $ip = $scaleObj }
			elseif ($null -ne $scaleObj)
			{
				if (($scaleObj.PSObject.Properties.Name -contains "ScaleBrand") -and ($scaleObj.ScaleBrand -match "Ishida"))
				{
					$isIshida = $true
				}
				if (($scaleObj.PSObject.Properties.Name -contains "ScaleBrand") -and ($scaleObj.ScaleBrand -match "BIZERBA"))
				{
					$isBizerba = $true
				}
				if ($scaleObj.PSObject.Properties.Name -contains "FullIP" -and $scaleObj.FullIP)
				{
					$ip = $scaleObj.FullIP
				}
				elseif (($scaleObj.PSObject.Properties.Name -contains "IPNetwork") -and ($scaleObj.PSObject.Properties.Name -contains "IPDevice") -and $scaleObj.IPNetwork -and $scaleObj.IPDevice)
				{
					$ip = (([string]$scaleObj.IPNetwork).TrimEnd('.') + '.' + [string]$scaleObj.IPDevice)
				}
			}
			if ($ip)
			{
				if ($isIshida)
				{
					Write_Log "Ishida scale [$ip] will use default VNC password.`r`n" "yellow"
					$AllVNCPasswords[$ip] = $DefaultVNCPassword
					continue
				}
				if ($isBizerba) { $BizerbaScales += @{ Host = $ip; Obj = $scaleObj }; continue }
				$NodeList += $ip
			}
		}
	}
	if ($BackofficeNumToMachineName) { $NodeList += $BackofficeNumToMachineName.Values | Where-Object { $_ } }
	$NodeList = $NodeList | Sort-Object -Unique

	if (($NodeList.Count -eq 0) -and ($BizerbaScales.Count -eq 0)) { throw "No machines provided for password extraction." }

	# 2. Ping regular nodes first (v2.6.9: parallel probe with a concurrency
	# cap - the old serial Test-Connection loop could block for minutes when
	# several nodes were down)
	$onlineProbe = @{}
	if ($NodeList.Count -gt 0) { $onlineProbe = Test-Hosts_Parallel -Targets $NodeList }
	$OnlineNodes = @()
	foreach ($node in $NodeList)
	{
		if ($onlineProbe[$node] -eq 'OK')
		{
			$OnlineNodes += $node
		}
		else
		{
			$AllVNCPasswords[$node] = $null
		}
	}

	# 3. Read ultravnc.ini of every online node in parallel - runspace threads, not one whole
	#    PowerShell process per node. The admin share first (fast, what works in the stores);
	#    WinRM only when the share cannot be read. (Before v2.7.7 WinRM came first and always
	#    failed: Invoke-Command has no -TimeoutSec in Windows PowerShell 5.1.)
	$readIniText = {
		param ($machineName, $uvncFolders)
		$findPasswd = {
			param ($lines)
			foreach ($line in $lines) { if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)') { return $matches[1] } }
			return $null
		}
		$password = $null
		foreach ($folder in $uvncFolders)
		{
			$iniPath = "\\$machineName\$folder\ultravnc.ini"
			try { if ([System.IO.File]::Exists($iniPath)) { $password = & $findPasswd ([System.IO.File]::ReadAllLines($iniPath)) } }
			catch { }
			if ($password) { break }
		}
		if (-not $password)
		{
			try
			{
				$opt = New-PSSessionOption -OpenTimeout 10000 -OperationTimeout 20000
				$lines = Invoke-Command -ComputerName $machineName -SessionOption $opt -ErrorAction Stop -ScriptBlock {
					param ($dirs)
					foreach ($d in $dirs) { $p = Join-Path $d 'ultravnc.ini'; if (Test-Path -LiteralPath $p) { Get-Content -LiteralPath $p; break } }
				} -ArgumentList (, @($uvncFolders | ForEach-Object { 'C:\' + $_.Substring(3) }))
				$password = & $findPasswd $lines
			}
			catch { }
		}
		return @{ Machine = $machineName; Password = $password }
	}.ToString()

	$MaxConcurrentJobs = 20
	if ($OnlineNodes.Count -gt 0)
	{
		$poolSize = [Math]::Min($MaxConcurrentJobs, [Math]::Max(1, $OnlineNodes.Count))
		$pool = [RunspaceFactory]::CreateRunspacePool(1, $poolSize)
		$pool.Open()
		$tasks = @(foreach ($machineName in $OnlineNodes)
			{
				$ps = [PowerShell]::Create()
				$ps.RunspacePool = $pool
				$null = $ps.AddScript($readIniText).AddArgument($machineName).AddArgument($uvncFolders)
				[pscustomobject]@{ Machine = $machineName; PS = $ps; Handle = $ps.BeginInvoke() }
			})
		$deadline = (Get-Date).AddSeconds(120)
		$hung = $false
		foreach ($task in $tasks)
		{
			$left = [int][Math]::Max(0, ($deadline - (Get-Date)).TotalMilliseconds)
			if ($task.Handle.AsyncWaitHandle.WaitOne($left))
			{
				try
				{
					$r = @($task.PS.EndInvoke($task.Handle)) | Select-Object -Last 1
					if ($r -and $r.Machine) { $AllVNCPasswords[$r.Machine] = $r.Password }
				}
				catch { $AllVNCPasswords[$task.Machine] = $null }
				$task.PS.Dispose()
			}
			else
			{
				# Still stuck after 2 minutes: give up on it without blocking on the stop
				$AllVNCPasswords[$task.Machine] = $null
				try { $null = $task.PS.BeginStop($null, $null) }
				catch { }
				$hung = $true
			}
		}
		if ($hung) { try { $null = $pool.BeginClose($null, $null) } catch { } }
		else { try { $pool.Close(); $pool.Dispose() } catch { } }
	}
	# 4. Handle Bizerba scales (using cmdkey/bizuser logic, serially)
	foreach ($b in $BizerbaScales)
	{
		$scaleHost = $b.Host
		$uvncIniRelativePaths = @(
			"Program Files\uvnc bvba\UltraVNC\ultravnc.ini",
			"Program Files (x86)\uvnc bvba\UltraVNC\ultravnc.ini"
		)
		$passwords = @("bizerba", "biyerba")
		$username = "bizuser"
		$password = $null
		$fullIniPath = $null

		# Ping check first
		if (-not (Test-Connection -ComputerName $scaleHost -Count 1 -Quiet -ErrorAction SilentlyContinue))
		{
			$AllVNCPasswords[$scaleHost] = $null
			continue
		}

		foreach ($uvncIniRel in $uvncIniRelativePaths)
		{
			foreach ($pw in $passwords)
			{
				# Clear any lingering share session, but keep the persisted credential manager entry.
				try { $null = & net.exe use "\\$scaleHost\c$" '/delete' '/y' 2>$null }
				catch { }
				cmdkey /add:$scaleHost /user:$username /pass:$pw | Out-Null
				$shareIniPath = "\\$scaleHost\c$\$uvncIniRel"
				if (Test-Path $shareIniPath -ErrorAction SilentlyContinue)
				{
					try
					{
						$content = Get-Content $shareIniPath -ErrorAction Stop
						foreach ($line in $content)
						{
							if ($line -match '^\s*passwd\s*=\s*([0-9A-Fa-f]+)')
							{
								$password = $matches[1]
								$fullIniPath = $shareIniPath
								break
							}
						}
					}
					catch { }
				}
				# Drop any temporary SMB session, but do not delete the saved credential.
				try { $null = & net.exe use "\\$scaleHost\c$" '/delete' '/y' 2>$null }
				catch { }
				if ($password) { break }
			}
			if ($password) { break }
		}
		$AllVNCPasswords[$scaleHost] = $password
	}
	$script:FunctionResults['AllVNCPasswords'] = $AllVNCPasswords
	return $AllVNCPasswords
}

# ===================================================================================================
#                              FUNCTION: Export_VNC_Files_For_All_Nodes
# ---------------------------------------------------------------------------------------------------
# Description:
#   Generates UltraVNC (.vnc) connection files for every lane, scale, and backoffice discovered
#   in the store, saved under Desktop\Lanes, Desktop\Scales, Desktop\Backoffices. Default VNC
#   password is 4330df922eb03b6e (UltraVNC encrypted form) unless a per-machine password is
#   available (a node that is off keeps the password already in its file); Ishida scale files
#   allow untrusted servers. Scales with missing IP info are skipped. Unchanged files are not
#   rewritten.
#
# Parameters:
#   - LaneNumToMachineName       [hashtable] (required): LaneNumber => MachineName
#   - ScaleCodeToIPInfo          [hashtable] (required): ScaleCode => scale object (with IP)
#   - BackofficeNumToMachineName [hashtable] (required): BackofficeTerminal => MachineName
#   - AllVNCPasswords            [hashtable]: MachineName => VNC password
#   - RunInBackground            [switch]: run as a background job
#   - SilentBackground           [switch]: suppress log window output
#   - InternalBackgroundJob      [switch]: internal flag - already inside a background job
# ===================================================================================================

function Export_VNC_Files_For_All_Nodes
{
	param (
		[Parameter(Mandatory = $true)]
		[hashtable]$LaneNumToMachineName,
		[Parameter(Mandatory = $true)]
		[hashtable]$ScaleCodeToIPInfo,
		[Parameter(Mandatory = $true)]
		[hashtable]$BackofficeNumToMachineName,
		[Parameter(Mandatory = $false)]
		[hashtable]$AllVNCPasswords,
		[Parameter(Mandatory = $false)]
		[switch]$RunInBackground,
		[Parameter(Mandatory = $false)]
		[switch]$SilentBackground,
		[Parameter(Mandatory = $false)]
		[switch]$InternalBackgroundJob
	)

	$log = {
		param (
			[string]$Message,
			[string]$Color = "Black"
		)

		if (-not $SilentBackground)
		{
			Write_Log $Message $Color
		}
	}

	$validLaneMap = @{ }
	if ($LaneNumToMachineName)
	{
		foreach ($kv in $LaneNumToMachineName.GetEnumerator())
		{
			$machineName = [string]$kv.Value
			if (-not [string]::IsNullOrWhiteSpace($machineName))
			{
				$validLaneMap[[string]$kv.Key] = $machineName.Trim()
			}
		}
	}

	$validScaleMap = @{ }
	if ($ScaleCodeToIPInfo)
	{
		foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
		{
			$scaleObj = $kv.Value
			$ip = $null
			if ($null -ne $scaleObj)
			{
				if (($scaleObj.PSObject.Properties.Name -contains 'FullIP') -and -not [string]::IsNullOrWhiteSpace([string]$scaleObj.FullIP))
				{
					$ip = [string]$scaleObj.FullIP
				}
				elseif (
					($scaleObj.PSObject.Properties.Name -contains 'IPNetwork') -and
					($scaleObj.PSObject.Properties.Name -contains 'IPDevice') -and
					-not [string]::IsNullOrWhiteSpace([string]$scaleObj.IPNetwork) -and
					-not [string]::IsNullOrWhiteSpace([string]$scaleObj.IPDevice)
				)
				{
					$ip = (([string]$scaleObj.IPNetwork).TrimEnd('.') + '.' + [string]$scaleObj.IPDevice)
				}
			}

			if (-not [string]::IsNullOrWhiteSpace($ip))
			{
				$validScaleMap[[string]$kv.Key] = $scaleObj
			}
		}
	}

	$validBackofficeMap = @{ }
	if ($BackofficeNumToMachineName)
	{
		foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
		{
			$machineName = [string]$kv.Value
			if (-not [string]::IsNullOrWhiteSpace($machineName))
			{
				$validBackofficeMap[[string]$kv.Key] = $machineName.Trim()
			}
		}
	}

	$LaneNumToMachineName = $validLaneMap
	$ScaleCodeToIPInfo = $validScaleMap
	$BackofficeNumToMachineName = $validBackofficeMap

	if ($RunInBackground -and -not $InternalBackgroundJob)
	{
		$hasAnyNodes = (
			($LaneNumToMachineName -and $LaneNumToMachineName.Count -gt 0) -or
			($ScaleCodeToIPInfo -and $ScaleCodeToIPInfo.Count -gt 0) -or
			($BackofficeNumToMachineName -and $BackofficeNumToMachineName.Count -gt 0)
		)

		if (-not $hasAnyNodes) { return }

		if (
			$script:VncExportTask -and
			$script:VncExportTask.Job -and
			$script:VncExportTask.Job.State -in @('Running', 'NotStarted')
		)
		{
			return
		}

		try
		{
			$writeLogDefinition = 'param([string]$Message, $Color, [switch]$IncludeTimestamp)'
			$getAllVncPasswordsDefinition = ${function:Get_All_VNC_Passwords}.ToString()
			$exportVncDefinition = ${function:Export_VNC_Files_For_All_Nodes}.ToString()
			# Get_All_VNC_Passwords probes the nodes with Test-Hosts_Parallel - the job needs it too
			# (missing before v2.7.0: the background export failed whenever it had to look up passwords)
			$testHostsDefinition = ${function:Test-Hosts_Parallel}.ToString()

			$job = Start-Job -Name 'VNCExportJob' -ScriptBlock {
				param (
					[string]$WriteLogDefinition,
					[string]$GetAllVncPasswordsDefinition,
					[string]$ExportVncDefinition,
					[string]$TestHostsDefinition,
					[hashtable]$LaneMap,
					[hashtable]$ScaleMap,
					[hashtable]$BackofficeMap,
					[hashtable]$AllVncPasswords,
					[string]$UltraVncFolder,
					[int]$AnsiCodePage
				)

				Add-Type -AssemblyName System.Windows.Forms 2>$null
				Add-Type -AssemblyName System.Drawing 2>$null

				Set-Item -Path function:Write_Log -Value ([scriptblock]::Create($WriteLogDefinition))
				Set-Item -Path function:Get_All_VNC_Passwords -Value ([scriptblock]::Create($GetAllVncPasswordsDefinition))
				Set-Item -Path function:Export_VNC_Files_For_All_Nodes -Value ([scriptblock]::Create($ExportVncDefinition))
				Set-Item -Path function:Test-Hosts_Parallel -Value ([scriptblock]::Create($TestHostsDefinition))

				$script:UltraVncFolder = $UltraVncFolder
				$script:ansiPcEncoding = [System.Text.Encoding]::GetEncoding($AnsiCodePage)
				$script:FunctionResults = @{ }

				try
				{
					Export_VNC_Files_For_All_Nodes `
												   -LaneNumToMachineName $LaneMap `
												   -ScaleCodeToIPInfo $ScaleMap `
												   -BackofficeNumToMachineName $BackofficeMap `
												   -AllVNCPasswords $AllVncPasswords `
												   -SilentBackground `
												   -InternalBackgroundJob

					$passwordCache = $null
					if ($script:FunctionResults.ContainsKey('AllVNCPasswords'))
					{
						$passwordCache = $script:FunctionResults['AllVNCPasswords']
					}

					[pscustomobject]@{
						Success		    = $true
						AllVNCPasswords = $passwordCache
						Error		    = $null
					}
				}
				catch
				{
					[pscustomobject]@{
						Success		    = $false
						AllVNCPasswords = $null
						Error		    = $_.Exception.Message
					}
				}
			} -ArgumentList @(
				$writeLogDefinition,
				$getAllVncPasswordsDefinition,
				$exportVncDefinition,
				$testHostsDefinition,
				$LaneNumToMachineName,
				$ScaleCodeToIPInfo,
				$BackofficeNumToMachineName,
				$AllVNCPasswords,
				$script:UltraVncFolder,
				1252
			)

			$script:VncExportTask = @{
				Job	    = $job
				Started = Get-Date
			}

			if (-not $script:VncExportReaper)
			{
				$script:VncExportReaper = New-Object System.Windows.Forms.Timer
				$script:VncExportReaper.Interval = 1000
				$script:VncExportReaper.Add_Tick({
						if (-not $script:VncExportTask)
						{
							try
							{
								if ($script:VncExportReaper)
								{
									$script:VncExportReaper.Stop()
									$script:VncExportReaper.Dispose()
									$script:VncExportReaper = $null
								}
							}
							catch { }
							return
						}

						$taskState = $script:VncExportTask
						if (-not $taskState.Job -or $taskState.Job.State -in @('Running', 'NotStarted')) { return }

						$result = $null
						try { $result = Receive-Job -Job $taskState.Job -ErrorAction SilentlyContinue }
						catch { }
						try { Remove-Job -Job $taskState.Job -Force -ErrorAction SilentlyContinue | Out-Null }
						catch { }

						$resultObj = $null
						if ($result)
						{
							$resultObj = @($result) | Select-Object -Last 1
						}

						if (
							$resultObj -and
							$resultObj.Success -and
							$resultObj.AllVNCPasswords -and
							($resultObj.AllVNCPasswords -is [hashtable]) -and
							$resultObj.AllVNCPasswords.Count -gt 0
						)
						{
							$script:FunctionResults['AllVNCPasswords'] = $resultObj.AllVNCPasswords
						}

						$script:VncExportTask = $null

						try
						{
							if ($script:VncExportReaper)
							{
								$script:VncExportReaper.Stop()
								$script:VncExportReaper.Dispose()
								$script:VncExportReaper = $null
							}
						}
						catch { }
					})
			}

			$script:VncExportReaper.Start()
		}
		catch
		{
			try { if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } }
			catch { }
			$script:VncExportTask = $null
		}

		return
	}

	& $log "`r`n==================== Starting Export_VNCFiles_ForAllNodes ====================`r`n" "blue"
	$DefaultVNCPassword = "4330df922eb03b6e"
	$desktop = [Environment]::GetFolderPath("Desktop")
	$lanesDir = Join-Path $desktop "Lanes"
	$scalesDir = Join-Path $desktop "Scales"
	$backofficesDir = Join-Path $desktop "BackOffices"
	$hasLaneNodes = ($LaneNumToMachineName -and $LaneNumToMachineName.Count -gt 0)
	$hasScaleNodes = ($ScaleCodeToIPInfo -and $ScaleCodeToIPInfo.Count -gt 0)
	$hasBackofficeNodes = ($BackofficeNumToMachineName -and $BackofficeNumToMachineName.Count -gt 0)

	if (-not ($hasLaneNodes -or $hasScaleNodes -or $hasBackofficeNodes))
	{
		& $log "No lanes, scales, or backoffices were found. Nothing to export." "yellow"
		& $log "`r`n==================== Export_VNCFiles_ForAllNodes Completed ====================" "blue"
		return
	}

	# ---- If passwords not provided, scan them ----
	if (-not $AllVNCPasswords -or $AllVNCPasswords.Count -eq 0)
	{
		& $log "Gathering VNC passwords for all machines`r`n" "magenta"
		$AllVNCPasswords = Get_All_VNC_Passwords -LaneNumToMachineName $LaneNumToMachineName -ScaleCodeToIPInfo $ScaleCodeToIPInfo -BackofficeNumToMachineName $BackofficeNumToMachineName
	}

	# --- Shared VNC file content with token ---
	$vncTemplate = @"
[connection]
host=%%HOST%%
port=5900
proxyhost=
proxyport=0
password=%%PASSWORD%%
[options]
use_encoding_0=1
use_encoding_1=1
use_encoding_2=1
use_encoding_3=0
use_encoding_4=1
use_encoding_5=1
use_encoding_6=1
use_encoding_7=1
use_encoding_8=1
use_encoding_9=1
use_encoding_10=1
use_encoding_11=0
use_encoding_12=0
use_encoding_13=0
use_encoding_14=0
use_encoding_15=0
use_encoding_16=1
use_encoding_17=1
use_encoding_18=1
use_encoding_19=1
use_encoding_20=0
use_encoding_21=0
use_encoding_22=0
use_encoding_23=0
use_encoding_24=0
use_encoding_25=1
use_encoding_26=1
use_encoding_27=1
use_encoding_28=0
use_encoding_29=1
preferred_encoding=10
restricted=0
AllowUntrustedServers=0
viewonly=0
nostatus=0
nohotkeys=0
showtoolbar=1
fullscreen=0
SavePos=0
SaveSize=0
GNOME=0
directx=0
autoDetect=0
8bit=0
shared=1
swapmouse=0
belldeiconify=0
BlockSameMouse=0
emulate3=1
JapKeyboard=0
emulate3timeout=100
emulate3fuzz=4
disableclipboard=0
localcursor=1
Scaling=0
AutoScaling=1
AutoScalingEven=0
AutoScalingLimit=0
scale_num=1
scale_den=1
cursorshape=1
noremotecursor=0
compresslevel=6
quality=8
ServerScale=1
Reconnect=3
EnableCache=0
EnableZstd=1
QuickOption=1
UseDSMPlugin=0
UseProxy=0
sponsor=0
allowMonitorSpanning=0
ChangeServerRes=0
extendDisplay=0
showExtend=0
use_virt=0
useAllMonitors=0
requestedWidth=0
requestedHeight=0
DSMPlugin=NoPlugin
folder=$script:UltraVncFolder
prefix=vnc_
imageFormat=.jpeg
InfoMsg=
AutoReconnect=3
ExitCheck=0
FileTransferTimeout=30
ListenPort=5500
KeepAliveInterval=5
ThrottleMouse=0
AutoAcceptIncoming=0
AutoAcceptNoDSM=0
RequireEncryption=0
PreemptiveUpdates=0
"@

	$resolveVncPassword = {
		param (
			[string]$LookupKey
		)

		if (
			$AllVNCPasswords -and
			-not [string]::IsNullOrWhiteSpace($LookupKey) -and
			$AllVNCPasswords.ContainsKey($LookupKey) -and
			-not [string]::IsNullOrWhiteSpace([string]$AllVNCPasswords[$LookupKey])
		)
		{
			return [string]$AllVNCPasswords[$LookupKey]
		}

		return $null   # unknown: writeExports keeps the existing file's password, else the default
	}

	$writeExports = {
		param (
			[System.Collections.ArrayList]$Exports,
			[string]$TargetDir,
			[string]$CategoryLabel
		)

		if (-not $Exports -or $Exports.Count -eq 0)
		{
			& $log "No $CategoryLabel VNC files to export. Skipping folder creation.`r`n" "gray"
			return 0
		}

		[void][System.IO.Directory]::CreateDirectory($TargetDir)

		$written = 0
		foreach ($item in $Exports)
		{
			$existing = $null
			if ([System.IO.File]::Exists($item.FilePath))
			{
				try { $existing = [System.IO.File]::ReadAllText($item.FilePath, $script:ansiPcEncoding) }
				catch { $existing = $null }
			}
			# A node that was off (password not read this time) keeps the password already in its
			# file - before, its real password was overwritten with the default
			$password = $item.Password
			if (-not $password -and $existing -and $existing -match '(?m)^password=([0-9A-Fa-f]+)\s*$') { $password = $matches[1] }
			if (-not $password) { $password = $DefaultVNCPassword }

			$content = $vncTemplate.Replace('%%HOST%%', $item.Host).Replace('%%PASSWORD%%', $password)
			if ($item.AllowUntrustedServers)
			{
				$content = $content -replace 'AllowUntrustedServers=0', 'AllowUntrustedServers=1'
			}

			# Unchanged files are left alone (no rewrite + desktop refresh at every launch)
			if ($existing -ceq $content) { continue }
			[System.IO.File]::WriteAllText($item.FilePath, $content, $script:ansiPcEncoding)
			& $log "Created: $($item.FilePath)" "green"
			$written++
		}

		& $log "$($Exports.Count) $CategoryLabel VNC files in $TargetDir ($written written, $($Exports.Count - $written) already current)`r`n" "blue"
		return $Exports.Count
	}

	# ---- Lanes ---- #
	& $log "-------------------- Exporting Lane VNC Files --------------------" "blue"
	$laneExports = New-Object System.Collections.ArrayList
	$dedupedLanes = @{ }
	foreach ($kv in $LaneNumToMachineName.GetEnumerator())
	{
		$machineName = [string]$kv.Value
		if ([string]::IsNullOrWhiteSpace($machineName)) { continue }

		# Only write one file per machine name
		$machineKey = $machineName.Trim().ToLowerInvariant()
		if (-not $dedupedLanes.ContainsKey($machineKey))
		{
			$dedupedLanes[$machineKey] = $true
			if ($machineName.ToUpper() -match '^(POS|SCO)\d+$')
			{
				$fileName = "$($machineName.ToUpper()).vnc"
			}
			else
			{
				$fileName = "Lane_$($kv.Key).vnc"
			}

			[void]$laneExports.Add([pscustomobject]@{
					FilePath = [System.IO.Path]::Combine($lanesDir, $fileName)
					Host	 = $machineName
					Password = (& $resolveVncPassword $machineName)
					AllowUntrustedServers = $false
				})
		}
	}
	$laneCount = & $writeExports $laneExports $lanesDir 'lane'

	# ---- Scales ---- #
	& $log "-------------------- Exporting Scale VNC Files --------------------" "blue"
	$scaleExports = New-Object System.Collections.ArrayList
	$dedupedScales = @{ }
	foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
	{
		$scaleObj = $kv.Value
		# Make sure to dedupe by IP address
		$ip = if ($scaleObj.FullIP) { $scaleObj.FullIP }
		elseif ($scaleObj.IPNetwork -and $scaleObj.IPDevice) { (([string]$scaleObj.IPNetwork).TrimEnd('.') + '.' + [string]$scaleObj.IPDevice) }
		else { $null }
		if ($ip)
		{
			$ipKey = $ip.Trim().ToLowerInvariant()
			if ($dedupedScales.ContainsKey($ipKey)) { continue }
			$dedupedScales[$ipKey] = $true

			$octets = $ip -split '\.'
			$lastOctet = $octets[-1]
			$brandRaw = ($scaleObj.ScaleBrand -as [string]).Trim()
			$model = ($scaleObj.ScaleModel -as [string]).Trim()
			$brand = if ($brandRaw)
			{
				($brandRaw -split ' ' | ForEach-Object {
						if ($_.Length -gt 0) { $_.Substring(0, 1).ToUpper() + $_.Substring(1).ToLower() }
						else { $_ }
					}) -join ' '
			}
			else { "" }
			if ($brand -and $model)
			{
				$fileName = "$brand($model)_${lastOctet}.vnc"
			}
			elseif ($brand)
			{
				$fileName = "$brand(Unknown)_${lastOctet}.vnc"
			}
			else
			{
				$fileName = "Scale_${lastOctet}.vnc"
			}

			[void]$scaleExports.Add([pscustomobject]@{
					FilePath = [System.IO.Path]::Combine($scalesDir, $fileName)
					Host	 = $ip
					Password = (& $resolveVncPassword $ip)
					AllowUntrustedServers = ($brand -like '*Ishida*')
				})
		}
	}
	$scaleCount = & $writeExports $scaleExports $scalesDir 'scale'

	# ---- Backoffices ---- #
	& $log "-------------------- Exporting Backoffice VNC Files --------------------" "blue"
	$boExports = New-Object System.Collections.ArrayList
	$dedupedBOs = @{ }
	foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
	{
		$boName = [string]$kv.Value
		$terminal = $kv.Key
		if ([string]::IsNullOrWhiteSpace($boName)) { continue }

		$boKey = $boName.Trim().ToLowerInvariant()
		if (-not $dedupedBOs.ContainsKey($boKey))
		{
			$dedupedBOs[$boKey] = $true
			$fileName = "Backoffice_${terminal}.vnc"

			[void]$boExports.Add([pscustomobject]@{
					FilePath = [System.IO.Path]::Combine($backofficesDir, $fileName)
					Host	 = $boName
					Password = (& $resolveVncPassword $boName)
					AllowUntrustedServers = $false
				})
		}
	}
	$boCount = & $writeExports $boExports $backofficesDir 'backoffice'

	& $log "VNC file export complete!" "green"
	& $log "`r`n==================== Export_VNCFiles_ForAllNodes Completed ====================" "blue"
}

# ===================================================================================================
#                              FUNCTION: Get_Remote_Machine_Info
# ---------------------------------------------------------------------------------------------------
# Description:
#   Collects hardware/software info from selected lanes, scales and backoffices: one runspace
#   thread per node (10 at a time) tries WMI, then CIM, then REG.EXE, each bounded by a timeout, with an
#   INI-based fallback for lanes/backoffices. Writes per-group info files and populates
#   $script:LaneHardwareInfo / $script:ScaleHardwareInfo / $script:BackofficeHardwareInfo.
#
# Parameters:
#   - RunInBackground       [switch]: run as a background job
#   - SilentBackground      [switch]: background mode without box logging
#   - InternalBackgroundJob [switch]: called from inside the background job itself
#   - UseAllAvailableNodes  [switch]: probe every discovered node (skip the selection form)
#   - SelectedLanes / SelectedScales / SelectedBackoffices [object[]]: pre-selected nodes
#
# Variables:
#   - $script:FunctionResults: lane/scale/backoffice maps and StoreNumber (input)
#
# Writes:
#   - Info text files under Desktop\Lanes, Desktop\Scales and Desktop\BackOffices
#
# Notes:
#   - Requires admin rights on the remote machines for WMI/CIM/REG/SC operations.
# ===================================================================================================

function Get_Remote_Machine_Info
{
	param (
		[switch]$RunInBackground,
		[switch]$SilentBackground,
		[switch]$InternalBackgroundJob,
		[switch]$UseAllAvailableNodes,
		[object[]]$SelectedLanes,
		[object[]]$SelectedScales,
		[object[]]$SelectedBackoffices
	)

	$log = {
		param (
			[string]$Message,
			$Color
		)
		if (-not $SilentBackground)
		{
			Write_Log $Message $Color
		}
	}

	& $log "`r`n==================== Starting Get_Remote_Machine_Info ====================`r`n" "blue"

	# --------------------------- Tunables ---------------------------
	$maxConcurrentJobs = 10
	$wmiTimeoutSeconds = 5
	$cimTimeoutSeconds = 10
	$usePSRemotingFallback = $false # set $true to try PS-Session registry read before REG.EXE

	# --------------------------- Outputs (reset) ---------------------------
	$script:LaneHardwareInfo = $null
	$script:ScaleHardwareInfo = $null
	$script:BackofficeHardwareInfo = $null

	# --------------------------- Inputs from your environment ---------------------------
	$LaneNumToMachineName = $script:FunctionResults['LaneNumToMachineName']
	$ScaleCodeToIPInfo = $script:FunctionResults['ScaleCodeToIPInfo']
	$BackofficeNumToMachineName = $script:FunctionResults['BackofficeNumToMachineName']
	$StoreNumber = $script:FunctionResults['StoreNumber']
	$DbsPath = $script:DbsPath

	# --------------------------- Keep only discovered/usable nodes ---------------------------
	$availableLaneNodes = @{ }
	if ($LaneNumToMachineName)
	{
		foreach ($kv in $LaneNumToMachineName.GetEnumerator())
		{
			$machineName = [string]$kv.Value
			if (-not [string]::IsNullOrWhiteSpace($machineName))
			{
				$availableLaneNodes[[string]$kv.Key] = $machineName.Trim()
			}
		}
	}

	$availableScaleNodes = @{ }
	if ($ScaleCodeToIPInfo)
	{
		foreach ($kv in $ScaleCodeToIPInfo.GetEnumerator())
		{
			$scaleObj = $kv.Value
			$scaleIp = $null
			if ($null -ne $scaleObj)
			{
				if (($scaleObj.PSObject.Properties.Name -contains 'FullIP') -and -not [string]::IsNullOrWhiteSpace([string]$scaleObj.FullIP))
				{
					$scaleIp = [string]$scaleObj.FullIP
				}
				elseif (
					($scaleObj.PSObject.Properties.Name -contains 'IPNetwork') -and
					($scaleObj.PSObject.Properties.Name -contains 'IPDevice') -and
					-not [string]::IsNullOrWhiteSpace([string]$scaleObj.IPNetwork) -and
					-not [string]::IsNullOrWhiteSpace([string]$scaleObj.IPDevice)
				)
				{
					$scaleIp = (([string]$scaleObj.IPNetwork).TrimEnd('.') + '.' + [string]$scaleObj.IPDevice)
				}
			}

			if (-not [string]::IsNullOrWhiteSpace($scaleIp))
			{
				$availableScaleNodes[[string]$kv.Key] = $scaleObj
			}
		}
	}

	$availableBackofficeNodes = @{ }
	if ($BackofficeNumToMachineName)
	{
		foreach ($kv in $BackofficeNumToMachineName.GetEnumerator())
		{
			$machineName = [string]$kv.Value
			if (-not [string]::IsNullOrWhiteSpace($machineName))
			{
				$availableBackofficeNodes[[string]$kv.Key] = $machineName.Trim()
			}
		}
	}

	$LaneNumToMachineName = $availableLaneNodes
	$ScaleCodeToIPInfo = $availableScaleNodes
	$BackofficeNumToMachineName = $availableBackofficeNodes

	$selectedLanes = @($SelectedLanes)
	$selectedScales = @($SelectedScales)
	$selectedBOs = @($SelectedBackoffices)

	if ($UseAllAvailableNodes -and
		(-not $selectedLanes -or $selectedLanes.Count -eq 0) -and
		(-not $selectedScales -or $selectedScales.Count -eq 0) -and
		(-not $selectedBOs -or $selectedBOs.Count -eq 0))
	{
		$selectedLanes = @()
		$selectedScales = @()
		$selectedBOs = @()

		if ($LaneNumToMachineName) { $selectedLanes = @($LaneNumToMachineName.Keys | Sort-Object) }
		if ($ScaleCodeToIPInfo) { $selectedScales = @($ScaleCodeToIPInfo.Keys | Sort-Object) }
		if ($BackofficeNumToMachineName) { $selectedBOs = @($BackofficeNumToMachineName.Keys | Sort-Object) }
	}

	if ((-not $UseAllAvailableNodes) -and
		(-not $selectedLanes -or $selectedLanes.Count -eq 0) -and
		(-not $selectedScales -or $selectedScales.Count -eq 0) -and
		(-not $selectedBOs -or $selectedBOs.Count -eq 0))
	{
		# --------------------------- Let user pick nodes ---------------------------
		$nodeSelection = Show_Node_Selection_Form -StoreNumber $StoreNumber `
												  -NodeTypes @("Lane", "Scale", "Backoffice") `
												  -Title "Select Nodes to Pull Hardware Info"

		if (-not $nodeSelection)
		{
			& $log "Get_Remote_Machine_Info cancelled by user." "yellow"
			& $log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
			return $false
		}

		$selectedLanes = @($nodeSelection.Lanes)
		$selectedScales = @($nodeSelection.Scales)
		$selectedBOs = @($nodeSelection.Backoffices)
	}

	if ((-not $selectedLanes -or $selectedLanes.Count -eq 0) -and
		(-not $selectedScales -or $selectedScales.Count -eq 0) -and
		(-not $selectedBOs -or $selectedBOs.Count -eq 0))
	{
		& $log "No nodes selected. Operation aborted." "yellow"
		& $log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
		return $false
	}

	if ($RunInBackground -and -not $InternalBackgroundJob)
	{
		if (
			$script:RemoteMachineInfoTask -and
			$script:RemoteMachineInfoTask.Job -and
			$script:RemoteMachineInfoTask.Job.State -in @('Running', 'NotStarted')
		)
		{
			& $log "Remote machine info export is already running in the background." "yellow"
			return $true
		}

		$job = $null
		try
		{
			$getRemoteMachineInfoDefinition = ${function:Get_Remote_Machine_Info}.ToString()
			$writeLogDefinition = 'param([string]$Message, $Color, [switch]$IncludeTimestamp)'

			$job = Start-Job -Name 'RemoteMachineInfoJob' -ScriptBlock {
				param (
					[string]$GetRemoteMachineInfoDefinition,
					[string]$WriteLogDefinition,
					[hashtable]$LaneMap,
					[hashtable]$ScaleMap,
					[hashtable]$BackofficeMap,
					[string]$StoreNumber,
					[string]$DbsPath,
					[bool]$SilentBackgroundMode,
					[object[]]$SelectedLanes,
					[object[]]$SelectedScales,
					[object[]]$SelectedBackoffices
				)

				Set-Item -Path function:Write_Log -Value ([scriptblock]::Create($WriteLogDefinition))
				Set-Item -Path function:Get_Remote_Machine_Info -Value ([scriptblock]::Create($GetRemoteMachineInfoDefinition))

				$script:FunctionResults = @{
					LaneNumToMachineName	   = $LaneMap
					ScaleCodeToIPInfo		   = $ScaleMap
					BackofficeNumToMachineName = $BackofficeMap
					StoreNumber			       = $StoreNumber
				}
				$script:DbsPath = $DbsPath

				try
				{
					Get_Remote_Machine_Info `
											-SelectedLanes $SelectedLanes `
											-SelectedScales $SelectedScales `
											-SelectedBackoffices $SelectedBackoffices `
											-InternalBackgroundJob `
											-SilentBackground:$SilentBackgroundMode | Out-Null

					[pscustomobject]@{
						Success			       = $true
						LaneHardwareInfo	   = $script:LaneHardwareInfo
						ScaleHardwareInfo	   = $script:ScaleHardwareInfo
						BackofficeHardwareInfo = $script:BackofficeHardwareInfo
						Error				   = $null
					}
				}
				catch
				{
					[pscustomobject]@{
						Success			       = $false
						LaneHardwareInfo	   = $null
						ScaleHardwareInfo	   = $null
						BackofficeHardwareInfo = $null
						Error				   = $_.Exception.Message
					}
				}
			} -ArgumentList @(
				$getRemoteMachineInfoDefinition,
				$writeLogDefinition,
				$LaneNumToMachineName,
				$ScaleCodeToIPInfo,
				$BackofficeNumToMachineName,
				$StoreNumber,
				$DbsPath,
				([bool]$SilentBackground),
				$selectedLanes,
				$selectedScales,
				$selectedBOs
			)

			$script:RemoteMachineInfoTask = @{
				Job	    = $job
				Started = Get-Date
				Silent  = [bool]$SilentBackground
			}

			if (-not $script:RemoteMachineInfoReaper)
			{
				$script:RemoteMachineInfoReaper = New-Object System.Windows.Forms.Timer
				$script:RemoteMachineInfoReaper.Interval = 1000
				$script:RemoteMachineInfoReaper.Add_Tick({
						if (-not $script:RemoteMachineInfoTask)
						{
							try
							{
								if ($script:RemoteMachineInfoReaper)
								{
									$script:RemoteMachineInfoReaper.Stop()
									$script:RemoteMachineInfoReaper.Dispose()
									$script:RemoteMachineInfoReaper = $null
								}
							}
							catch { }
							return
						}

						$taskState = $script:RemoteMachineInfoTask
						if (-not $taskState.Job -or $taskState.Job.State -in @('Running', 'NotStarted')) { return }

						$result = $null
						try { $result = Receive-Job -Job $taskState.Job -ErrorAction SilentlyContinue }
						catch { }
						try { Remove-Job -Job $taskState.Job -Force -ErrorAction SilentlyContinue | Out-Null }
						catch { }

						$resultObj = $null
						if ($result)
						{
							$resultObj = @($result) | Select-Object -Last 1
						}

						if ($resultObj -and $resultObj.Success)
						{
							$script:LaneHardwareInfo = $resultObj.LaneHardwareInfo
							$script:ScaleHardwareInfo = $resultObj.ScaleHardwareInfo
							$script:BackofficeHardwareInfo = $resultObj.BackofficeHardwareInfo
							if (-not $taskState.Silent)
							{
								Write_Log "Remote machine info export completed in background." "green"
							}
						}
						elseif (-not $taskState.Silent)
						{
							$errText = if ($resultObj -and $resultObj.Error) { $resultObj.Error }
							else { "Unknown error." }
							Write_Log "Remote machine info export failed in background: $errText" "red"
						}

						$script:RemoteMachineInfoTask = $null

						try
						{
							if ($script:RemoteMachineInfoReaper)
							{
								$script:RemoteMachineInfoReaper.Stop()
								$script:RemoteMachineInfoReaper.Dispose()
								$script:RemoteMachineInfoReaper = $null
							}
						}
						catch { }
					})
			}

			if (-not $SilentBackground)
			{
				& $log "Remote machine info export started in background." "yellow"
			}

			$script:RemoteMachineInfoReaper.Start()
		}
		catch
		{
			try { if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } }
			catch { }
			$script:RemoteMachineInfoTask = $null
			& $log "Failed to start remote machine info export in background: $_" "red"
			return $false
		}

		return $true
	}


	# --------------------------- Per-node probe (runs in a runspace pool) ---------------------------
	# 1) ping, 2) WMI = CIM over DCOM, then CIM over WinRM - each with an operation timeout instead
	# of the old nested Start-Job per protocol (three PowerShell processes per node), 3) optional
	# PS-remoting registry read, 4) REG.EXE, 5) INI augmentation.
	$nodeProbeText = {
		param (
			$NodeName,
			$NodeNumber,
			$iniPath,
			$wmiTimeoutSeconds,
			$cimTimeoutSeconds,
			$usePSRemotingFallback
		)

		# --- 0) Result object scaffold ---
		$info = [PSCustomObject]@{
			Success			    = $false
			SystemManufacturer  = $null
			SystemProductName   = $null
			CPU				    = $null
			RAM				    = $null
			OSInfo			    = $null
			BIOS			    = $null
			Method			    = $null
			Error			    = $null
			MachineNameOverride = $null
		}

		# Win32 classes through one CIM session; $null when unreachable / no data
		$cimProbe = {
			param ([string]$Protocol, [int]$TimeoutSec)
			$session = $null
			try
			{
				$session = New-CimSession -ComputerName $NodeName -SessionOption (New-CimSessionOption -Protocol $Protocol) -OperationTimeoutSec $TimeoutSec -ErrorAction Stop
				$q = @{ CimSession = $session; OperationTimeoutSec = $TimeoutSec; ErrorAction = 'Stop' }
				$sys = Get-CimInstance @q -ClassName Win32_ComputerSystem
				if (-not ($sys -and $sys.Manufacturer -and $sys.Model)) { return $null }
				$cpu = Get-CimInstance @q -ClassName Win32_Processor | Select-Object -First 1
				$os = Get-CimInstance @q -ClassName Win32_OperatingSystem

				$biosVerOut = $null
				try
				{
					$bios = Get-CimInstance @q -ClassName Win32_BIOS
					if ($bios)
					{
						if ($bios.SMBIOSBIOSVersion) { $biosVerOut = $bios.SMBIOSBIOSVersion }
						elseif ($bios.BIOSVersion) { $biosVerOut = ($bios.BIOSVersion | Where-Object { $_ } | Select-Object -First 1) }
						elseif ($bios.Version) { $biosVerOut = $bios.Version }
						$dt = $null
						if ($bios.ReleaseDate -is [datetime]) { $dt = $bios.ReleaseDate }
						elseif ($bios.ReleaseDate) { try { $dt = [System.Management.ManagementDateTimeConverter]::ToDateTime([string]$bios.ReleaseDate) } catch { } }
						if ($dt)
						{
							if ($biosVerOut) { $biosVerOut = "$biosVerOut ($($dt.ToString('yyyy-MM-dd')))" }
							else { $biosVerOut = $dt.ToString('yyyy-MM-dd') }
						}
					}
				}
				catch { }

				[PSCustomObject]@{
					SystemManufacturer = $sys.Manufacturer
					SystemProductName  = $sys.Model
					CPU			       = $cpu.Name
					RAM			       = [math]::Round($sys.TotalPhysicalMemory / 1GB, 1)
					OSInfo			   = "$($os.Caption) ($($os.Version))"
					BIOS			   = $biosVerOut
					Method			   = $(if ($Protocol -eq 'Dcom') { 'WMI' } else { 'CIM' })
				}
			}
			catch { $null }
			finally { if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue } }
		}
		$applyCim = {
			param ($r)
			$info.SystemManufacturer = $r.SystemManufacturer
			$info.SystemProductName = $r.SystemProductName
			$info.CPU = $r.CPU
			$info.RAM = $r.RAM
			$info.OSInfo = $r.OSInfo
			if ($r.BIOS) { $info.BIOS = $r.BIOS }
			$info.Method = $r.Method
			$info.Success = $true
			$info.Error = $null
		}

		# --- 1) Ping (1.5 s; Test-Connection waited up to 4 s per switched-off node) ---
		$online = $false
		try
		{
			$pinger = New-Object System.Net.NetworkInformation.Ping
			try { $online = ($pinger.Send($NodeName, 1500).Status -eq [System.Net.NetworkInformation.IPStatus]::Success) }
			finally { $pinger.Dispose() }
		}
		catch { $online = $false }

		if (-not $online)
		{
			$info.Error = "Offline or unreachable."
			$info.Method = "Offline"
		}
		else
		{
			# --- 2) WMI (DCOM), then CIM (WinRM) ---
			$r = & $cimProbe 'Dcom' $wmiTimeoutSeconds
			if ($r) { & $applyCim $r }
			else
			{
				$info.Error = "WMI failed (access issue or null)"
				$r = & $cimProbe 'Wsman' $cimTimeoutSeconds
				if ($r) { & $applyCim $r }
				else { $info.Error = "CIM failed (access issue or null)" }
			}
		}
		# --- 3) PS-Remoting Registry (optional middle step) ---
		if (-not $info.Success -and $usePSRemotingFallback)
		{
			$regResult = $null
			try
			{
				$sessionOption = New-PSSessionOption -ConnectionTimeout 15000 -OpenTimeout 15000
				$session = New-PSSession -ComputerName $NodeName -SessionOption $sessionOption -ErrorAction SilentlyContinue
				if ($session)
				{
					$regResult = Invoke-Command -Session $session -ScriptBlock {
						$out = [PSCustomObject]@{
							SystemManufacturer = $null
							SystemProductName  = $null
							CPU			       = $null
							OSInfo			   = $null
							BIOS			   = $null
							Success		       = $false
							Error			   = $null
							Method			   = 'REG(PSRemoting)'
						}
						try
						{
							$manuf = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemManufacturer -ErrorAction SilentlyContinue
							$prod = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemProductName -ErrorAction SilentlyContinue
							$cpu = Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0' -Name ProcessorNameString -ErrorAction SilentlyContinue
							$os = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue

							# BIOS pieces
							$biosVendor = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSVendor -ErrorAction SilentlyContinue).BIOSVendor
							$biosVerStr = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSVersion -ErrorAction SilentlyContinue).BIOSVersion
							if (-not $biosVerStr)
							{
								$biosVerStr = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name SystemBiosVersion -ErrorAction SilentlyContinue).SystemBiosVersion
							}
							$biosDate = (Get-ItemProperty 'HKLM:\HARDWARE\DESCRIPTION\System\BIOS' -Name BIOSReleaseDate -ErrorAction SilentlyContinue).BIOSReleaseDate

							# Build OSInfo
							$osInfo = $null
							if ($os -and $os.ProductName) { $osInfo = $os.ProductName }
							if ($os -and $os.DisplayVersion)
							{
								if ($osInfo) { $osInfo = "$osInfo ($($os.DisplayVersion))" }
								else { $osInfo = $os.DisplayVersion }
							}
							elseif ($os -and $os.CurrentBuild)
							{
								if ($osInfo) { $osInfo = "$osInfo (Build $($os.CurrentBuild))" }
								else { $osInfo = "Build $($os.CurrentBuild)" }
							}

							# BIOS assemble
							$biosOut = $null
							if ($biosVendor -and $biosVerStr) { $biosOut = "$biosVendor $biosVerStr" }
							elseif ($biosVerStr) { $biosOut = $biosVerStr }
							elseif ($biosVendor) { $biosOut = $biosVendor }
							if ($biosDate)
							{
								if ($biosOut) { $biosOut = "$biosOut ($biosDate)" }
								else { $biosOut = $biosDate }
							}

							$out.SystemManufacturer = $manuf.SystemManufacturer
							$out.SystemProductName = $prod.SystemProductName
							$out.CPU = $cpu.ProcessorNameString
							$out.OSInfo = $osInfo
							$out.BIOS = $biosOut
							$out.Success = $true
						}
						catch { $out.Error = $_.Exception.Message }
						return $out
					} 2>$null
					Remove-PSSession $session -ErrorAction SilentlyContinue
				}
			}
			catch
			{
				if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
				$regResult = [PSCustomObject]@{
					SystemManufacturer = $null
					SystemProductName  = $null
					CPU			       = $null
					OSInfo			   = $null
					BIOS			   = $null
					Success		       = $false
					Error			   = "PS-Remoting registry query failed"
					Method			   = 'REG(PSRemoting)'
				}
			}

			if ($regResult -and $regResult.Success)
			{
				if (-not $info.SystemManufacturer) { $info.SystemManufacturer = $regResult.SystemManufacturer }
				if (-not $info.SystemProductName) { $info.SystemProductName = $regResult.SystemProductName }
				if (-not $info.CPU) { $info.CPU = $regResult.CPU }
				if (-not $info.OSInfo) { $info.OSInfo = $regResult.OSInfo }
				if (-not $info.BIOS -and $regResult.BIOS) { $info.BIOS = $regResult.BIOS }
				$info.Method = $regResult.Method
				$info.Success = $true
				$info.Error = $null
			}
			elseif (-not $info.Error)
			{
				if ($regResult) { $info.Error = $regResult.Error; $info.Method = 'REG(PSRemoting)' }
			}
		}

		# --- 4) REG.EXE (starts RemoteRegistry if needed, then restores) ---
		if (-not $info.Success)
		{
			$wasRunning = $false
			$wasDisabled = $false
			try
			{
				$state = sc.exe "\\$NodeName" query RemoteRegistry 2>$null | Select-String 'STATE'
				$start = sc.exe "\\$NodeName" qc    RemoteRegistry 2>$null | Select-String 'START_TYPE'
				if ($state -and $state.Line -match 'RUNNING') { $wasRunning = $true }
				if ($start -and $start.Line -match 'DISABLED') { $wasDisabled = $true }
				if ($wasDisabled) { sc.exe "\\$NodeName" config RemoteRegistry start= demand 2>$null | Out-Null }
				if (-not $wasRunning) { sc.exe "\\$NodeName" start  RemoteRegistry 2>$null | Out-Null }

				# tiny inline getter: returns only the DATA (no "REG_SZ")
				$getVal = {
					param ($hive,
						$path,
						$name)
					$patternName = [regex]::Escape($name)
					$raw = reg.exe QUERY "\\$NodeName\$hive\$path" /v $name 2>$null
					if (-not $raw) { return $null }
					$line = $raw | Select-String -Pattern "^\s*$patternName\s+REG_\w+\s+.+$" | Select-Object -First 1
					if (-not $line) { return $null }
					$data = $line.Line -replace "^\s*$patternName\s+REG_\w+\s+", ""
					$data.Trim()
				}

				$manuf = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS'               'SystemManufacturer'
				$prod = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS'               'SystemProductName'
				$cpu = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\CentralProcessor\0' 'ProcessorNameString'
				$osN = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'ProductName'
				$osB = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'DisplayVersion'
				$osV = & $getVal 'HKLM' 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'   'CurrentBuild'

				# BIOS from REG.EXE
				$biosVendor = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSVendor'
				$biosVerStr = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSVersion'
				if (-not $biosVerStr)
				{
					$biosVerStr = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'SystemBiosVersion'
				}
				$biosDate = & $getVal 'HKLM' 'HARDWARE\DESCRIPTION\System\BIOS' 'BIOSReleaseDate'
				if ($biosVerStr)
				{
					$biosVerStr = $biosVerStr -replace '\x00', ' '
					$biosVerStr = ($biosVerStr -split '\s{2,}' | Where-Object { $_ }) -join ' '
					$biosVerStr = $biosVerStr.Trim()
				}
				$biosOut = $null
				if ($biosVendor -and $biosVerStr) { $biosOut = "$biosVendor $biosVerStr" }
				elseif ($biosVerStr) { $biosOut = $biosVerStr }
				elseif ($biosVendor) { $biosOut = $biosVendor }
				if ($biosDate)
				{
					if ($biosOut) { $biosOut = "$biosOut ($biosDate)" }
					else { $biosOut = $biosDate }
				}

				# If CurrentBuild came as "0x... (7601)" keep the number in parentheses
				if ($osV -and $osV -match '\((\d+)\)') { $osV = $matches[1] }

				# build a clean OS string
				$osInfo = $null
				if ($osN) { $osInfo = $osN }
				if ($osB)
				{
					if ($osInfo) { $osInfo = "$osInfo ($osB)" }
					else { $osInfo = $osB }
				}
				elseif ($osV)
				{
					if ($osInfo) { $osInfo = "$osInfo (Build $osV)" }
					else { $osInfo = "Build $osV" }
				}

				if ($manuf -or $prod -or $cpu -or $osInfo -or $biosOut)
				{
					if (-not $info.SystemManufacturer) { $info.SystemManufacturer = $manuf }
					if (-not $info.SystemProductName) { $info.SystemProductName = $prod }
					if (-not $info.CPU) { $info.CPU = $cpu }
					if (-not $info.OSInfo) { $info.OSInfo = $osInfo }
					if (-not $info.BIOS -and $biosOut) { $info.BIOS = $biosOut }
					$info.Method = "REG.EXE"
					$info.Success = $true
					$info.Error = $null
				}
				else
				{
					$info.Method = "REG.EXE"
					$info.Success = $false
					$info.Error = "REG queries returned no data"
				}
			}
			catch
			{
				$info.Method = "REG.EXE"
				$info.Success = $false
				$info.Error = "REG queries failed"
			}
			finally
			{
				try
				{
					if (-not $wasRunning) { sc.exe "\\$NodeName" stop RemoteRegistry 2>$null | Out-Null }
					if ($wasDisabled) { sc.exe "\\$NodeName" config RemoteRegistry start= disabled 2>$null | Out-Null }
				}
				catch { }
			}
		}

		# A node that did not answer the ping and not REG.EXE either is reported as offline
		# (REG.EXE is still tried: some lanes block ICMP)
		if (-not $info.Success -and -not $online)
		{
			$info.Error = "Offline or unreachable."
			$info.Method = "Offline"
		}

		# --- 5) INI augmentation (optional; fills gaps if any fields still null) ---
		$returnInfo = [PSCustomObject]@{
			Machine    = $NodeName
			NodeNumber = $NodeNumber
			Info	   = $info
			IniFound   = $false
			IniPath    = $null
		}
		try
		{
			if ($iniPath)
			{
				$iniFolder = Split-Path $iniPath -Parent
				$iniLeaf = Split-Path $iniPath -Leaf
				$iniFile = Get-ChildItem -Path $iniFolder -Filter $iniLeaf -ErrorAction SilentlyContinue |
				Sort-Object LastWriteTime -Descending |
				Select-Object -First 1
				if ($iniFile -and (Test-Path $iniFile.FullName))
				{
					$returnInfo.IniFound = $true
					$returnInfo.IniPath = $iniFile.FullName
					$iniLines = Get-Content $iniFile.FullName -Encoding UTF8 -ErrorAction SilentlyContinue

					# tiny parser
					$sections = @{ }
					$curSec = ""
					foreach ($line in $iniLines)
					{
						if ($line -match '^\[(.+)\]$')
						{
							$curSec = $matches[1]; $sections[$curSec] = @{ }
						}
						elseif ($line -match '^\s*([^=]+?)\s*=\s*(.*)$' -and $curSec)
						{
							$sections[$curSec][$matches[1].Trim()] = $matches[2].Trim()
						}
					}
					if (-not $info.CPU -and $sections.ContainsKey('PROCESSOR') -and $sections['PROCESSOR'].ContainsKey('Cores'))
					{
						$cores = $sections['PROCESSOR']['Cores']
						$arch = $sections['PROCESSOR']['Architecture']
						if ($arch) { $info.CPU = "$cores cores ($arch)" }
						else { $info.CPU = "$cores cores" }
					}
					if ($info.RAM -eq $null -and $sections.ContainsKey('Memory') -and $sections['Memory'].ContainsKey('PhysicalMemory'))
					{
						$ramMb = $sections['Memory']['PhysicalMemory']
						if ($ramMb -match '^\d+$') { $info.RAM = [math]::Round([double]$ramMb/1024, 1) }
						else { $info.RAM = $ramMb }
					}
					if (-not $info.OSInfo -and $sections.ContainsKey('OperatingSystem') -and $sections['OperatingSystem'].ContainsKey('ProductName'))
					{
						$info.OSInfo = $sections['OperatingSystem']['ProductName']
					}
				}
			}
		}
		catch { }

		return $returnInfo

	}.ToString()

	# --------------------------- Output folders ---------------------------
	$desktop = [Environment]::GetFolderPath("Desktop")
	$lanesDir = Join-Path $desktop "Lanes"
	$scalesDir = Join-Path $desktop "Scales"
	$backofficesDir = Join-Path $desktop "BackOffices"

	# --------------------------- Process 3 categories uniformly ---------------------------
	foreach ($section in @(
			@{ Name = 'Lanes'; Selected = $selectedLanes; Dir = $lanesDir; FileName = 'Lanes_Info.txt' },
			@{ Name = 'Scales'; Selected = $selectedScales; Dir = $scalesDir; FileName = 'Scales_Info.txt' },
			@{ Name = 'BackOffices'; Selected = $selectedBOs; Dir = $backofficesDir; FileName = 'Backoffices_Info.txt' }
		))
	{
		if (-not $section.Selected -or $section.Selected.Count -eq 0) { continue }

		& $log "Processing $($section.Name) nodes..." "yellow"
		$results = @{ }
		$infoLines = New-Object System.Collections.Generic.List[string]

		# One node's result -> its info-file line + results entry
		$collectNode = {
			param ([string]$remoteName, $result)
			$info = if ($result) { $result.Info } else { $null }
			if (-not $info) { $info = [PSCustomObject]@{ Success = $false; Error = 'No answer within the time limit'; Method = 'Timeout'; BIOS = $null } }

			$line = "Machine Name: $remoteName |"
			if ($info.Success)
			{
				$line += " Manufacturer: $($info.SystemManufacturer) | Model: $($info.SystemProductName) | CPU: $($info.CPU)"
				if ($null -ne $info.RAM) { $line += " | RAM: $($info.RAM) GB" }
				if ($info.OSInfo) { $line += " | OS: $($info.OSInfo)" }
				& $log "Processed $remoteName ($($section.Name)): Success [$($info.Method)]" "green"
			}
			elseif ($info.Error)
			{
				$line += " [Hardware info unavailable] Error: $($info.Error)"
				& $log "Processed $remoteName ($($section.Name)): Error [$($info.Method)] - $($info.Error)" "red"
			}
			else
			{
				$line += " [No hardware info found]"
				& $log "Processed $remoteName ($($section.Name)): No info [$($info.Method)]" "yellow"
			}
			$biosText = if ($info.BIOS) { $info.BIOS } else { "Unknown" }
			$line += " | BIOS: $biosText | Method: $($info.Method)"
			$infoLines.Add($line)
			$results[$remoteName] = $info
		}

		$pool = [RunspaceFactory]::CreateRunspacePool(1, $maxConcurrentJobs)
		$pool.Open()
		$queued = @()
		foreach ($sel in $section.Selected)
		{
			# -------- Canonical per-node identity: $NodeNumber (code) + $NodeName (machine/IP) --------
			$NodeNumber = $null
			$NodeName = $null

			if ($section.Name -eq 'Lanes')
			{
				if ($sel.PSObject.Properties.Name -contains 'LaneNumber' -and $sel.LaneNumber)
				{
					$NodeNumber = "{0:D3}" -f [int]$sel.LaneNumber
				}
				else { $NodeNumber = "$sel" }
				$NodeName = $LaneNumToMachineName[$NodeNumber]
			}
			elseif ($section.Name -eq 'Scales')
			{
				$NodeNumber = "$sel"
				if ($ScaleCodeToIPInfo.ContainsKey($NodeNumber) -and
					$ScaleCodeToIPInfo[$NodeNumber].PSObject.Properties.Name -contains 'FullIP')
				{
					$NodeName = $ScaleCodeToIPInfo[$NodeNumber].FullIP
				}
			}
			elseif ($section.Name -eq 'BackOffices')
			{
				if ($sel.PSObject.Properties.Name -contains 'BONumber' -and $sel.BONumber)
				{
					$NodeNumber = "{0:D3}" -f [int]$sel.BONumber
				}
				else { $NodeNumber = "$sel" }
				$NodeName = $BackofficeNumToMachineName[$NodeNumber]
			}

			if (-not $NodeName)
			{
				& $log "Skipping $($section.Name) $NodeNumber (no NodeName/mapping found)" "red"
				continue
			}

			# -------- Optional INI path (used to augment fields) --------
			$iniPattern = "INFO_${StoreNumber}${NodeNumber}_SMSStart.ini"
			if ($section.Name -eq 'Lanes') { $iniPath = Join-Path "\\$NodeName\storeman\office\dbs" $iniPattern }
			elseif ($section.Name -eq 'BackOffices') { $iniPath = Join-Path $DbsPath $iniPattern }
			else { $iniPath = $null }

			$ps = [PowerShell]::Create()
			$ps.RunspacePool = $pool
			$null = $ps.AddScript($nodeProbeText).AddArgument($NodeName).AddArgument($NodeNumber).AddArgument($iniPath).AddArgument($wmiTimeoutSeconds).AddArgument($cimTimeoutSeconds).AddArgument($usePSRemotingFallback)
			$queued += [pscustomobject]@{ Name = [string]$NodeName; PS = $ps; Handle = $ps.BeginInvoke() }
		} # end foreach selected

		# Collect in order; 90 s for the first batch of nodes plus 30 s per further batch
		$deadline = (Get-Date).AddSeconds(90 + 30 * [Math]::Max(0, [Math]::Ceiling($queued.Count / $maxConcurrentJobs) - 1))
		$hung = $false
		foreach ($q in $queued)
		{
			$result = $null
			$left = [int][Math]::Max(0, ($deadline - (Get-Date)).TotalMilliseconds)
			if ($q.Handle.AsyncWaitHandle.WaitOne($left))
			{
				try { $result = @($q.PS.EndInvoke($q.Handle)) | Select-Object -Last 1 }
				catch { $result = $null }
				$q.PS.Dispose()
			}
			else
			{
				try { $null = $q.PS.BeginStop($null, $null) }
				catch { }
				$hung = $true
			}
			& $collectNode $q.Name $result
		}
		if ($hung) { try { $null = $pool.BeginClose($null, $null) } catch { } }
		else { try { $pool.Close(); $pool.Dispose() } catch { } }

		# Write file (stable sorting: IPs zero-padded; otherwise lexical)
		if ($infoLines.Count)
		{
			[void][System.IO.Directory]::CreateDirectory($section.Dir)

			$sortedLines = $infoLines | Sort-Object {
				if ($_ -match "^Machine Name: ([^|]+)")
				{
					$name = $matches[1].Trim()
					if ($name -match '^\d{1,3}(\.\d{1,3}){3}$')
					{
						$oct = $name -split '\.'
						'{0:D3}.{1:D3}.{2:D3}.{3:D3}' -f [int]$oct[0], [int]$oct[1], [int]$oct[2], [int]$oct[3]
					}
					else { $name }
				}
				else { $_ }
			}
			$filePath = Join-Path $section.Dir $section.FileName
			$sortedLines -join "`r`n" | Set-Content -Path $filePath -Encoding Default
			& $log "Exported $($section.Name) info to $filePath" "green"
		}

		# Publish programmatic map
		if ($section.Name -eq 'Lanes') { $script:LaneHardwareInfo = $results }
		elseif ($section.Name -eq 'Scales') { $script:ScaleHardwareInfo = $results }
		elseif ($section.Name -eq 'BackOffices') { $script:BackofficeHardwareInfo = $results }

		& $log "Completed processing $($section.Name).`r`n" "green"
	} # end sections

	& $log "==================== Get_Remote_Machine_Info Completed ====================" "blue"
}

# ===================================================================================================
#                                SECTION: Scales Maintenance tool
# ---------------------------------------------------------------------------------------------------
#   Store-server tool: a check list of the four scale operations, run in order with "Run selected".
#   The operations are the verbatim functions above; Edit_TBS_SCL_ver520_Table and
#   Remove_Duplicate_Files_From_toBizerba keep their own windows (grid editor / monitor options).
# ===================================================================================================

$script:ScalesSteps = @(
	@{ Name = 'Repair BMS service'; Function = 'Repair_BMS'; Default = $true
		Info = 'Kill BMS, clean toBizerba + terminals cache, re-register BMSSrv.exe (-reg) and start the BMS service.' },
	@{ Name = 'Remove duplicate files from toBizerba'; Function = 'Remove_Duplicate_Files_From_toBizerba'; Default = $false
		Info = 'Duplicate File Monitor for toBizerba: run now, or install / uninstall it as a Windows service (NSSM).' },
	@{ Name = 'Edit TBS_SCL_ver520 table'; Function = 'Edit_TBS_SCL_ver520_Table'; Default = $false
		Info = 'Grid editor for the scale table: edit, Auto Sort (BIZERBA first), Save.' },
	@{ Name = 'Fix DEPLOY_CHG.sql'; Function = 'Fix_Deploy_CHG'; Default = $true
		Info = 'Restore the ScaleManagementApp deploy line in Office\DEPLOY_CHG.sql and remove the minutes task.' },
	@{ Name = 'Insert test item'; Function = 'Insert_Test_Item'; Default = $false
		Info = 'Insert the deterministic scale test item (one of three test PLUs) into SCL_TAB, OBJ_TAB, POS_TAB, PRICE_TAB and SCL_TXT_TAB - all in one transaction.' },
	@{ Name = 'Troubleshoot ScaleCommApp'; Function = 'Troubleshoot_ScaleCommApp'; Default = $false
		Info = 'Audit + fix every ScaleCommApp *.exe.config (StoreName, first scale IP, SQL instance, database, Recs_BatchSendFull; .bak backups) and restore the DEPLOY_CHG.sql deploy line.' }
)

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Initialize_ScalesContext - the $script: paths the scale functions read (same values as
#   the old tool's configuration section), based on the detected storeman folder
# ---------------------------------------------------------------------------------------------------
function Initialize_ScalesContext
{
	$alex = $script:Cfg.AlexRoot
	$script:AlexRoot = $alex
	$script:AlexProgramDataRoot = Join-Path $env:ProgramData 'Tecnica_Systems\Alex_C.T'
	$script:ScriptsLog = ($script:Paths.Log.TrimEnd('\') + '\')
	$script:ToolsDir = Join-Path $alex 'Tools'
	$script:NssmDefaultPath = Join-Path $script:ToolsDir 'NSSM\nssm.exe'
	$script:DuplicateFileMonitorServiceRoot = Join-Path $script:AlexProgramDataRoot 'Duplicate_File_Monitor_Service'

	$root = if ($script:Db -and $script:Db.Storeman) { $script:Db.Storeman } else { 'C:\Storeman' }
	$script:BasePath = $root
	$script:OfficePath = Join-Path $root 'Office'

	# ScaleCommApp (Ishida / ScaleManagementApp)
	$script:ScaleCommRoots = @("C:\ScaleCommApp", "D:\ScaleCommApp")
	$script:PrimaryScaleCommRoot = $script:ScaleCommRoots[0]
	$script:ScaleManagementAppExeName = "ScaleManagementApp.exe"
	$script:ScaleManagementAppFastDeployExeName = "ScaleManagementApp_FastDEPLOY.exe"
	$script:ScaleManagementAppExePath = Join-Path $script:PrimaryScaleCommRoot $script:ScaleManagementAppExeName
	$script:ScaleManagementAppFastDeployPath = Join-Path $script:PrimaryScaleCommRoot $script:ScaleManagementAppFastDeployExeName
	$script:ScaleManagementAppPathCandidates = @($script:ScaleCommRoots | ForEach-Object { [System.IO.Path]::Combine($_, $script:ScaleManagementAppExeName) })
	$script:ScaleManagementAppFastDeployPathCandidates = @($script:ScaleCommRoots | ForEach-Object { [System.IO.Path]::Combine($_, $script:ScaleManagementAppFastDeployExeName) })
	$patterns = @(foreach ($r in $script:ScaleCommRoots)
		{
			[regex]::Escape([System.IO.Path]::Combine($r, $script:ScaleManagementAppExeName))
			[regex]::Escape([System.IO.Path]::Combine($r, $script:ScaleManagementAppFastDeployExeName))
		}) | Select-Object -Unique
	$script:ScaleCommDeployExecRegex = "(?i)@EXEC\(RUN='(?:" + (@($patterns) -join '|') + ")'\);"

	# Bizerba BMS / RetailConnect
	$script:BmsRootPath = "C:\Bizerba\RetailConnect\BMS"
	$script:BmssrvExePath = Join-Path $script:BmsRootPath "BMSSrv.exe"
	$script:BmsExePath = Join-Path $script:BmsRootPath "BMS.exe"
	$script:BmsToBizerbaPath = Join-Path $script:BmsRootPath "toBizerba"
	$script:BmsTerminalsPath = Join-Path $script:BmsRootPath "terminals"

	# Database context for Edit_TBS_SCL_ver520_Table (old tool: Get_Store_And_Database_Info)
	if (-not $script:FunctionResults) { $script:FunctionResults = @{ } }
	if ($script:Db -and $script:Db.Ok)
	{
		$script:FunctionResults['DBSERVER'] = $script:Db.Server
		$script:FunctionResults['DBNAME'] = $script:Db.Database
		$script:FunctionResults['ConnectionString'] = $script:Db.ConnectionString
	}
	if (Test-Path -LiteralPath $script:NssmDefaultPath) { $script:NssmExe = $script:NssmDefaultPath; $script:FunctionResults['NssmExe'] = $script:NssmDefaultPath }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Write_Log - the old tool's logger, as a shim: colour name -> terminal line (+ file log)
# ---------------------------------------------------------------------------------------------------
function Write_Log
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[AllowEmptyString()]
		[string]$Message,
		$Color = 'Black',
		[switch]$IncludeTimestamp,
		[switch]$Raw,
		[switch]$FileOnly
	)
	$text = $Message.Trim("`r", "`n")
	if (-not $text.Trim()) { return }
	# Inside a background job (Start_Background_Function_Job ships this body) there is no terminal
	if (-not (Get-Command -Name Write_Term -CommandType Function -ErrorAction SilentlyContinue)) { Write-Host $text; return }
	$kind = switch -regex (([string]$Color).Trim().ToLower()) { '^green$' { 'ok' } '^red$' { 'fail' } '^(yellow|orange)$' { 'warn' } '^(blue|cyan)$' { 'info' } '^(gray|grey)$' { 'muted' } default { 'text' } }
	if ($text -match '^\s*=+\s+(.+?)\s+=+\s*$') { $text = $Matches[1]; $kind = 'step' }
	if ($FileOnly) { Write_LogFile $text; return }
	Write_Term $text $kind
	[System.Windows.Forms.Application]::DoEvents()
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Show_Scales_Info - what is installed on this server (terminal)
# ---------------------------------------------------------------------------------------------------
function Show_Scales_Info
{
	Initialize_ScalesContext
	Write_Term 'Scales Maintenance - this server' 'step'
	$svc = Get-Service -Name 'BMS' -ErrorAction SilentlyContinue
	$bms = if (Test-Path -LiteralPath $script:BmssrvExePath) { 'installed' + $(if ($svc) { ", service $($svc.Status)" } else { ', service not registered' }) } else { "not found ($script:BmssrvExePath)" }
	Write_Term ('BMS                 : {0}' -f $bms) 'text' -NoTime
	$tb = if (Test-Path -LiteralPath $script:BmsToBizerbaPath) { '{0} file(s)' -f @(Get-ChildItem -LiteralPath $script:BmsToBizerbaPath -File -ErrorAction SilentlyContinue).Count } else { 'folder not found' }
	Write_Term ('toBizerba           : {0}' -f $tb) 'text' -NoTime
	$mon = Get-Service -Name 'Remove_Duplicate_Files_From_toBizerba' -ErrorAction SilentlyContinue
	Write_Term ('Duplicate monitor   : {0}' -f $(if ($mon) { "service $($mon.Status)" } else { 'not installed' })) 'text' -NoTime
	$deploy = Join-Path $script:OfficePath 'DEPLOY_CHG.sql'
	$dText = 'not found'
	if (Test-Path -LiteralPath $deploy)
	{
		$hasLine = [bool](Select-String -LiteralPath $deploy -Pattern $script:ScaleCommDeployExecRegex -Quiet -ErrorAction SilentlyContinue)
		$dText = if ($hasLine) { 'deploy line present' } else { 'deploy line MISSING' }
	}
	Write_Term ('DEPLOY_CHG.sql      : {0}' -f $dText) $(if ($dText -like '*MISSING') { 'warn' } else { 'text' }) -NoTime
	$app = @($script:ScaleManagementAppFastDeployPathCandidates + $script:ScaleManagementAppPathCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)
	Write_Term ('ScaleManagementApp  : {0}' -f $(if ($app) { $app[0] } else { 'not found in C:\ / D:\ScaleCommApp' })) 'text' -NoTime
	$scaleCount = $null
	if ($script:Db -and $script:Db.Ok)
	{
		$cn = New-Object System.Data.SqlClient.SqlConnection $script:Db.ConnectionString
		try
		{
			$cn.Open(); $cmd = $cn.CreateCommand()
			$cmd.CommandText = "IF OBJECT_ID('TBS_SCL_ver520','U') IS NOT NULL SELECT COUNT(*) FROM TBS_SCL_ver520 ELSE SELECT -1"
			$scaleCount = [int]$cmd.ExecuteScalar()
		}
		catch { }
		finally { $cn.Dispose() }
	}
	$sText = if ($null -eq $scaleCount) { 'database not reachable' } elseif ($scaleCount -lt 0) { 'table not found' } else { "$scaleCount scale(s)" }
	Write_Term ('TBS_SCL_ver520      : {0}' -f $sText) 'text' -NoTime
	Write_Term ('NSSM                : {0}' -f $(if (Test-Path -LiteralPath $script:NssmDefaultPath) { $script:NssmDefaultPath } else { 'not downloaded yet (fetched when the duplicate monitor needs it)' })) 'text' -NoTime
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Invoke_ScalesStep - one scales operation (Start_OpsRun calls it); $true when it worked
# ---------------------------------------------------------------------------------------------------
function Invoke_ScalesStep
{
	param ($Step)
	if ($Step.Function -eq 'Remove_Duplicate_Files_From_toBizerba') { Confirm_Nssm 'install the monitor as a service' }
	switch ($Step.Function)
	{
		'Fix_Deploy_CHG' { return [bool](Fix_Deploy_CHG -OfficePath $script:OfficePath | Select-Object -Last 1) }
		'Insert_Test_Item'
		{
			if (-not ($script:Db -and $script:Db.Ok)) { Write_Term 'Not connected to the store database.' 'fail'; return $false }
			$null = Insert_Test_Item -ConnectionString $script:Db.ConnectionString
			return $true
		}
		'Troubleshoot_ScaleCommApp'
		{
			# store name as the old tool resolves it (system.ini NAME=)
			try { Get_Store_Name_From_INI -INIPath (Join-Path $script:OfficePath 'system.ini') } catch { }
			$null = Troubleshoot_ScaleCommApp -AllMatches -Force -OfficePath $script:OfficePath
			return $true
		}
		default { $null = & $Step.Function; return $true }
	}
}

# NSSM is fetched (old tool: at every start) only when a service install may need it
function Confirm_Nssm
{
	param ([string]$Why)
	if (Test-Path -LiteralPath $script:NssmDefaultPath) { return }
	Write_Term ("NSSM (needed to $Why) is not on this server yet - downloading...") 'info'
	[System.Windows.Forms.Application]::DoEvents()
	$dl = Download_NSSM -Wait -WaitTimeoutSeconds 90
	if (Test-Path -LiteralPath $script:NssmDefaultPath)
	{
		Write_Term "NSSM ready: $script:NssmDefaultPath" 'ok'
		$script:NssmExe = $script:NssmDefaultPath
		$script:FunctionResults['NssmExe'] = $script:NssmDefaultPath
	}
	else { Write_Term ('NSSM download failed{0} - "Run Now" still works, installing the service will not.' -f $(if ($dl -and $dl.Error) { ": $($dl.Error)" } else { '' })) 'warn' }
}

# ===================================================================================================
#                                  SECTION: Remote Lane Tools
# ---------------------------------------------------------------------------------------------------
#   Store-server tool: INI Editor, Remove Archive Bit, Delete DBS, Refresh PIN pad files, Close open
#   transactions, Send Restart All Programs - the verbatim functions above, on lanes picked in the
#   old tool's lane picker. Lanes come from Retrieve_Nodes (store database TER_TAB / Ter_Load.sql).
# ===================================================================================================

$script:RemoteSteps = @(
	@{ Name = 'INI Editor'; Function = 'INI_Editor'; Default = $false
		Info = 'Edit an INI (Setup.ini, Setting.ini, System.ini, ApiVerifoneMX.ini...) of the server or a lane, save it, then copy / merge it to other lanes. Backups go to Backups\INI_Editor.' },
	@{ Name = 'Remove Archive Bit'; Function = 'Remove_Archive_Bit'; Default = $false
		Info = 'Clear the archive bit so the Storeman load re-copies changed files: run once now, or install / uninstall the always-on watcher service (NSSM).' },
	@{ Name = 'Delete DBS files on lanes'; Function = 'Delete_DBS'; Default = $false
		Info = 'Delete *.txt and *.dwr (optionally *.sus) from \\lane\Storeman\Office\DBS on the lanes you pick.' },
	@{ Name = 'Refresh PIN pad files'; Function = 'Refresh_PIN_Pad_Files'; Default = $false
		Info = 'Touch PreferredAIDs.ini, EMVCAPKey.ini and EMVAID.ini on the picked lanes so the PIN pads reload them; then offers Restart All Programs.' },
	@{ Name = 'Close open transactions'; Function = 'Close_Open_Transactions'; Default = $false
		Info = 'Close the transactions reported in the XE health files (or all open transactions on the lanes you pick) via SQL or a Close_Transaction.sqi drop, then restart the lane programs.' },
	@{ Name = 'Send Restart All Programs'; Function = 'Send_Restart_All_Programs'; Default = $false
		Info = 'Send @exec(RESTART_ALL=PROGRAMS) to the SMSStart mailslot of the lanes you pick.' },
	@{ Name = 'Clear XE folder'; Function = 'Clear_XE_Folder'; Default = $false
		Info = 'Clean Office\XE<store>901 (urgent messages): keep current health files, delete the rest, then keep sweeping every 30 s while the toolkit is open (new FATAL files are kept).' },
	@{ Name = 'Send server time to lanes'; Function = 'Send_SERVER_time_to_Lanes'; Default = $false
		Info = 'Set the server time on the lanes you pick - once now, or as a recurring schedule (you are asked which).' },
	@{ Name = 'Send SMS key to lanes'; Function = 'Send_SMS_Key_to_Lanes'; Default = $false
		Info = 'Sync the server SMS key to the lane databases: Windows auth, then SQL auth (Tecnica), then an SQI file; can also install the recurring sync task.' }
)

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Initialize_RemoteContext - the old tool's $script: paths + store/lane context
# ---------------------------------------------------------------------------------------------------
function Initialize_RemoteContext
{
	Initialize_ScalesContext
	$alex = $script:Cfg.AlexRoot
	$script:LoadPath = Join-Path $script:OfficePath 'Load'
	$script:SetupFiles = Join-Path $alex 'Setup_Files'
	$script:BackupRoot = ((Join-Path $alex 'Backups').TrimEnd('\') + '\')
	$script:ScriptsFolder = ((Join-Path $alex 'Scripts').TrimEnd('\') + '\')
	$script:StoremanRoots = @('C:\Storeman', 'D:\Storeman')
	foreach ($d in @($script:SetupFiles, $script:BackupRoot, $script:ScriptsFolder))
	{
		if (-not (Test-Path -LiteralPath $d)) { try { $null = New-Item -Path $d -ItemType Directory -Force } catch { } }
	}
	if (-not $script:LaneProtocols) { $script:LaneProtocols = @{ } }
	$script:FunctionResults['StoreNumber'] = if ($script:Db -and $script:Db.Store) { $script:Db.Store } else { 'N/A' }
	$script:FunctionResults['StoremanPath'] = $script:BasePath
	if (-not $script:FunctionResults.ContainsKey('SqlModuleName')) { $script:FunctionResults['SqlModuleName'] = 'None' }

	# more of the old tool's configuration, read by the time sync / SMS key / export functions
	$script:StoreNumber = $script:FunctionResults['StoreNumber']
	$script:PrimaryStoremanRoot = $script:BasePath
	$script:StoremanOfficeRoots = @(@($script:BasePath) + $script:StoremanRoots | Select-Object -Unique | ForEach-Object { [System.IO.Path]::Combine($_, 'Office') })
	$script:PrimaryStoremanOfficePath = $script:OfficePath
	$script:ProtocolResultsFile = Join-Path $script:SetupFiles 'Protocol_Results.txt'
	$script:BackupSqlUser = $script:Cfg.SqlUser
	$script:BackupSqlPass = $script:Cfg.SqlPass
	$script:UltraVncFolder = 'C:\Users\Administrator\Documents\UltraVNC'
	$script:ansiPcEncoding = [System.Text.Encoding]::GetEncoding(1252)
	$script:DbsPath = Join-Path $script:OfficePath 'Dbs'
	$script:IsTestEnvironment = $false
}

# Lanes of the store (Retrieve_Nodes, once per session or on Refresh)
function Confirm_RemoteNodes
{
	param ([switch]$Refresh)
	if (-not $Refresh -and $script:FunctionResults['LaneNumToMachineName'] -and $script:FunctionResults['LaneNumToMachineName'].Count) { return $true }
	$store = $script:FunctionResults['StoreNumber']
	if (-not $store -or $store -eq 'N/A') { Write_Term 'Store number unknown (Startup.ini STORE=) - cannot list the lanes.' 'fail'; return $false }
	Set_Status 'Finding the lanes of the store...'
	[System.Windows.Forms.Application]::DoEvents()
	$nodes = $null
	try { $nodes = Retrieve_Nodes -StoreNumber $store -RefreshCache:$Refresh }
	catch { Write_Term ('Lane discovery failed: {0}' -f $_.Exception.Message) 'fail' }
	Set_Status 'Remote Lane Tools'
	return [bool]($script:FunctionResults['LaneNumToMachineName'] -and $script:FunctionResults['LaneNumToMachineName'].Count)
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Show_Remote_Info - store + lanes (terminal)
# ---------------------------------------------------------------------------------------------------
function Show_Remote_Info
{
	param ([switch]$Refresh)
	Initialize_RemoteContext
	Write_Term 'Remote Lane Tools - store lanes' 'step'
	$store = $script:FunctionResults['StoreNumber']
	Write_Term ('Store #  : {0}' -f $store) 'text' -NoTime
	if (-not (Confirm_RemoteNodes -Refresh:$Refresh)) { Write_Term 'No lanes found (TER_TAB / Office\Load\Ter_Load.sql).' 'warn' -NoTime; return }
	$map = $script:FunctionResults['LaneNumToMachineName']
	$lanes = @($map.Keys | Where-Object { $_ -match '^\d{3}$' -and $_ -notmatch '^(8|9)' } | Sort-Object)
	Write_Term ('Lanes    : {0}' -f $lanes.Count) 'text' -NoTime
	$line = ''
	foreach ($ln in $lanes)
	{
		$item = '{0}={1}' -f $ln, $map[$ln]
		if (($line.Length + $item.Length) -gt 100) { Write_Term ('           ' + $line.TrimEnd(', ')) 'text' -NoTime; $line = '' }
		$line += $item + ', '
	}
	if ($line) { Write_Term ('           ' + $line.TrimEnd(', ')) 'text' -NoTime }
	$xe = Join-Path $script:OfficePath ('XE{0}901' -f $store)
	$xeCount = if (Test-Path -LiteralPath $xe) { @(Get-ChildItem -LiteralPath $xe -Filter 'S*.???' -File -ErrorAction SilentlyContinue).Count } else { $null }
	Write_Term ('XE folder: {0}' -f $(if ($null -eq $xeCount) { "not found ($xe)" } else { "$xeCount health report file(s)" })) 'text' -NoTime
	$svc = Get-Service -Name 'Remove_Archive_Bit' -ErrorAction SilentlyContinue
	Write_Term ('Archive-bit service: {0}' -f $(if ($svc) { $svc.Status } else { 'not installed' })) 'text' -NoTime
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Initialize_LaneProtocols - what Close_Open_Transactions needs from the old tool's
#   background protocol loop: the SQL module name + each lane's reachable protocol (TCP / Named
#   Pipes / File). One parallel probe, no scheduled tasks are created on the lanes.
# ---------------------------------------------------------------------------------------------------
function Initialize_LaneProtocols
{
	param ([switch]$AlwaysProbe)   # SMS key sync uses SqlClient: probe even without a SQL module
	if ($script:FunctionResults['SqlModuleName'] -eq 'None' -and -not $script:SqlModuleChecked)
	{
		$script:SqlModuleChecked = $true
		Set_Status 'Looking for the SQL PowerShell module...'
		[System.Windows.Forms.Application]::DoEvents()
		$m = Get-Module -ListAvailable -Name SqlServer, SQLPS -ErrorAction SilentlyContinue | Select-Object -First 1
		if ($m) { $script:FunctionResults['SqlModuleName'] = [string]$m.Name }
	}
	if ($script:FunctionResults['SqlModuleName'] -eq 'None' -and -not $AlwaysProbe)
	{
		Write_Term 'No SQL PowerShell module on this server - transactions are closed with Close_Transaction.sqi files (same as the old tool).' 'muted'
		return
	}
	$map = $script:FunctionResults['LaneNumToMachineName']
	$lanes = @($map.Keys | Where-Object { $_ -match '^\d{3}$' -and $_ -notmatch '^(8|9)' })
	if (-not $lanes) { return }
	Write_Term ('Checking the SQL connection of {0} lane(s)...' -f $lanes.Count) 'info'
	$pool = [runspacefactory]::CreateRunspacePool(1, 8)
	$pool.Open()
	$jobs = @()
	foreach ($ln in $lanes)
	{
		$info = $null
		try { $info = Get_All_Lanes_Database_Info -LaneNumber $ln } catch { }
		if (-not $info) { $script:LaneProtocols[$ln] = 'File'; continue }
		$ps = [PowerShell]::Create()
		$ps.RunspacePool = $pool
		$null = $ps.AddScript({
				param ($Tcp, $Np)
				foreach ($p in @(@('TCP', $Tcp), @('Named Pipes', $Np)))
				{
					if (-not $p[1]) { continue }
					$c = New-Object System.Data.SqlClient.SqlConnection (([string]$p[1]).TrimEnd(';') + ';Connect Timeout=3;')
					try { $c.Open(); return $p[0] } catch { } finally { $c.Dispose() }
				}
				'File'
			}).AddArgument($info['TcpConnStr']).AddArgument($info['NamedPipesConnStr'])
		$jobs += @{ Lane = $ln; PS = $ps; Handle = $ps.BeginInvoke() }
	}
	$deadline = [datetime]::Now.AddSeconds(20)
	while (@($jobs | Where-Object { -not $_.Handle.IsCompleted }).Count -and [datetime]::Now -lt $deadline)
	{
		[System.Windows.Forms.Application]::DoEvents()
		Start-Sleep -Milliseconds 50
	}
	foreach ($j in $jobs)
	{
		$proto = 'File'
		if ($j.Handle.IsCompleted) { try { $proto = [string](@($j.PS.EndInvoke($j.Handle)) | Select-Object -Last 1) } catch { } }
		$script:LaneProtocols[$j.Lane] = $proto
		$j.PS.Dispose()
	}
	$pool.Close()
	$summary = ($script:LaneProtocols.GetEnumerator() | Where-Object { $_.Key -match '^\d{3}$' } | Sort-Object Key | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
	Write_Term ('Lane SQL access: {0}' -f $summary) 'muted'
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Invoke_RemoteStep - one remote lane operation (Start_OpsRun calls it)
# ---------------------------------------------------------------------------------------------------
function Invoke_RemoteStep
{
	param ($Step)
	$store = $script:FunctionResults['StoreNumber']
	$needsLanes = $Step.Function -notin @('Remove_Archive_Bit', 'Clear_XE_Folder')
	if ($needsLanes -and -not (Confirm_RemoteNodes)) { Write_Term 'No lanes known for this store - skipped.' 'fail'; return $false }
	if ($Step.Function -in @('Send_Restart_All_Programs', 'Close_Open_Transactions', 'Refresh_PIN_Pad_Files', 'Send_SERVER_time_to_Lanes'))
	{
		try { Ensure_MailslotSender_Type }
		catch { Write_Term ('Could not load the mailslot sender: {0}' -f $_.Exception.Message) 'fail'; return $false }
	}
	switch ($Step.Function)
	{
		'INI_Editor' { $null = INI_Editor }
		'Remove_Archive_Bit' { Confirm_Nssm 'install the archive-bit service'; $null = Remove_Archive_Bit }
		'Close_Open_Transactions' { Initialize_LaneProtocols; $null = Close_Open_Transactions -StoreNumber $store }
		'Clear_XE_Folder'
		{
			$xe = Join-Path $script:OfficePath ('XE{0}901' -f $store)
			$job = Clear_XE_Folder -FolderPath $xe
			if ($job) { $script:ClearXEJob = $job; Write_Term 'XE watcher running in the background until the toolkit is closed.' 'info' }
			return [bool]$job
		}
		'Send_SERVER_time_to_Lanes'
		{
			$ans = Read_Prompt 'Schedule a recurring time sync? (Y = schedule, N = one-time now)' 'N' '^(?i)(y|yes|n|no)$' 'Y or N.'
			if ($null -eq $ans) { return $false }
			$null = Send_SERVER_time_to_Lanes -StoreNumber $store -Schedule:($ans -match '^(?i)y')
		}
		'Send_SMS_Key_to_Lanes' { Initialize_LaneProtocols -AlwaysProbe; $null = Send_SMS_Key_to_Lanes -StoreNumber $store }
		default { $null = & $Step.Function -StoreNumber $store }
	}
	return $true
}

# ===================================================================================================
#                                   SECTION: Repair Windows tool
# ---------------------------------------------------------------------------------------------------
#   Lane or server. Recommended order: DISM check -> DISM scan -> DISM repair -> component cleanup ->
#   SFC (after DISM, so SFC repairs from a healthy component store) -> check disk -> disk health ->
#   event log report. Optional: Windows Update reset (when DISM cannot download repair files) and a
#   full check disk at the next restart. Tool output streams live into the terminal.
# ===================================================================================================

$script:WindowsSteps = @(
	@{ Name = 'DISM - quick health check'; Function = 'DismCheck'; Default = $true
		Info = 'DISM /Online /Cleanup-Image /CheckHealth - seconds. Reports whether the Windows component store is already flagged as damaged.' },
	@{ Name = 'DISM - full scan'; Function = 'DismScan'; Default = $true
		Info = 'DISM /ScanHealth - 5 to 20 minutes. Scans the whole component store for damage.' },
	@{ Name = 'DISM - repair'; Function = 'DismRestore'; Default = $true
		Info = 'DISM /RestoreHealth - repairs the component store (downloads clean files from Windows Update). 10 to 30 minutes.' },
	@{ Name = 'DISM - component cleanup'; Function = 'DismCleanup'; Default = $true
		Info = 'DISM /StartComponentCleanup - removes superseded update files; frees space and shrinks the store.' },
	@{ Name = 'SFC - system file check'; Function = 'Sfc'; Default = $true
		Info = 'sfc /scannow - verifies and repairs protected Windows files, using the store DISM just repaired. 5 to 20 minutes.' },
	@{ Name = 'Check disk - online scan'; Function = 'ChkScan'; Default = $true
		Info = 'chkdsk X: /scan on every fixed NTFS drive - finds file system errors without a restart.' },
	@{ Name = 'Disk health report'; Function = 'DiskHealth'; Default = $true
		Info = 'Physical disks: health status, SMART failure prediction, temperature, wear, uncorrected read errors, power-on hours; free space per drive.' },
	@{ Name = 'Event Viewer - critical and error events'; Function = 'Events'; Default = $true
		Info = 'System + Application logs: critical and error events of the last N days, grouped by source and event ID, with description, codes and a hint for well-known events.' },
	@{ Name = 'Reset Windows Update components'; Function = 'WuReset'; Default = $false
		Info = 'Stops the update services, renames SoftwareDistribution and catroot2, starts them again. Use when DISM repair fails with 0x800F081F / 0x800F0906 or updates keep failing.' },
	@{ Name = 'Check disk C: with repair at next restart'; Function = 'ChkdskReboot'; Default = $false
		Info = 'Schedules chkdsk C: /f /r for the next restart (can take hours on a big drive). Works in any Windows language.' }
)

# System32 even from a 32-bit PowerShell on 64-bit Windows (DISM / SFC refuse to run through WOW64)
function Get_SystemToolPath
{
	param ([string]$Exe)
	$dir = if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) { 'Sysnative' } else { 'System32' }
	Join-Path (Join-Path $env:SystemRoot $dir) $Exe
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Invoke_LiveProcess - runs a console tool, streams its lines into the terminal (progress
#   lines with a % go to the status bar instead), keeps the window responsive. Returns Exit/Text.
# ---------------------------------------------------------------------------------------------------
function Invoke_LiveProcess
{
	param ([string]$Exe, [string]$Arguments, [string]$Encoding = 'oem', [string]$StdIn, [switch]$Quiet)
	$enc = if ($Encoding -eq 'unicode') { [System.Text.Encoding]::Unicode } else { [System.Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage) }
	$sync = [hashtable]::Synchronized(@{ Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'; Progress = ''; Done = $false; Exit = $null; Error = $null })
	$ps = [PowerShell]::Create()
	$null = $ps.AddScript({
			param ($Sync, $Exe, $Arguments, $Enc, $StdIn)
			try
			{
				$psi = New-Object System.Diagnostics.ProcessStartInfo $Exe, $Arguments
				$psi.UseShellExecute = $false
				$psi.CreateNoWindow = $true
				$psi.RedirectStandardOutput = $true
				$psi.RedirectStandardError = $true
				$psi.StandardOutputEncoding = $Enc
				$psi.StandardErrorEncoding = $Enc
				if ($StdIn) { $psi.RedirectStandardInput = $true }
				$p = [System.Diagnostics.Process]::Start($psi)
				if ($StdIn) { $p.StandardInput.WriteLine($StdIn); $p.StandardInput.Close() }
				$errTask = $p.StandardError.ReadToEndAsync()
				$sr = $p.StandardOutput
				$sb = New-Object System.Text.StringBuilder
				while (($c = $sr.Read()) -ge 0)
				{
					if ($c -eq 13 -or $c -eq 10)
					{
						if ($sb.Length)
						{
							$line = $sb.ToString(); [void]$sb.Clear()
							if ($line -match '\d{1,3}([.,]\d)?\s?%' -and $line -notmatch '100([.,]0)?\s?%') { $Sync.Progress = $line.Trim() } else { $Sync.Queue.Enqueue($line) }
						}
					}
					elseif ($c -ne 0) { [void]$sb.Append([char]$c) }
				}
				if ($sb.Length) { $Sync.Queue.Enqueue($sb.ToString()) }
				$p.WaitForExit()
				foreach ($l in ($errTask.Result -split "`r?`n")) { if ($l.Trim()) { $Sync.Queue.Enqueue($l) } }
				$Sync.Exit = $p.ExitCode
			}
			catch { $Sync.Error = $_.Exception.Message }
			$Sync.Done = $true
		}).AddArgument($sync).AddArgument($Exe).AddArgument($Arguments).AddArgument($enc).AddArgument($StdIn)
	$h = $ps.BeginInvoke()
	$t0 = [datetime]::Now
	$all = New-Object System.Collections.Generic.List[string]
	$tool = [System.IO.Path]::GetFileNameWithoutExtension($Exe).ToUpper()
	while ($true)
	{
		$line = $null
		while ($sync.Queue.TryDequeue([ref]$line))
		{
			$t = $line.Trim()
			if ($t) { $all.Add($t); if (-not $Quiet) { Write_Term ('    ' + $t) 'text' -NoTime } }
		}
		if ($sync.Done -and $sync.Queue.IsEmpty) { break }
		Set_Status ('{0} running ({1:mm\:ss})   {2}' -f $tool, ([datetime]::Now - $t0), $sync.Progress)
		[System.Windows.Forms.Application]::DoEvents()
		Start-Sleep -Milliseconds 40
	}
	try { $ps.EndInvoke($h) } catch { }
	$ps.Dispose()
	if ($sync.Error) { Write_Term ("Could not run {0}: {1}" -f $Exe, $sync.Error) 'fail' }
	[pscustomobject]@{ Exit = $sync.Exit; Text = ($all -join "`n"); Duration = ([datetime]::Now - $t0); Error = $sync.Error }
}

function Format_Hex { param ($Code) if ($null -eq $Code) { 'none' } else { '0x{0:X8}' -f [int]$Code } }

# DISM exit codes (same explanations as the old server maintenance worker)
function Write_DismResult
{
	param ([string]$What, $R)
	if ($R.Error) { return $false }
	$why = switch ($R.Exit)
	{
		-2146498529 { 'repair files not found - Windows Update is not reachable (try "Reset Windows Update components", or a repair source)' }   # 0x800F081F
		-2146498298 { 'could not download repair files - check internet / WSUS access' }    # 0x800F0906
		-2146498554 { 'changes are pending - restart the machine, then run it again' }      # 0x800F0806
		740 { 'needs administrator rights' }
		87 { 'option not supported on this Windows version' }
		default { 'exit code ' + (Format_Hex $R.Exit) }
	}
	$dur = '{0:mm\:ss}' -f $R.Duration
	if ($R.Exit -eq 0)
	{
		if ($R.Text -match '(?i)repairable|reparable') { Write_Term "$What - the component store is DAMAGED but repairable: run DISM repair ($dur)" 'warn'; return $true }
		Write_Term "$What - OK ($dur)" 'ok'; return $true
	}
	if ($R.Exit -eq 3010) { Write_Term "$What - done, restart the machine to finish ($dur)" 'warn'; return $true }
	Write_Term "$What - FAILED: $why ($dur)" 'fail'
	return $false
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Write_EventReport - critical + error events, grouped, readable
# ---------------------------------------------------------------------------------------------------
$script:EventHints = @{
	'Microsoft-Windows-Kernel-Power|41'          = 'Unexpected shutdown / power loss (no clean shutdown). Check UPS, power button presses, crashes.'
	'Microsoft-Windows-WER-SystemErrorReporting|1001' = 'Blue screen (bugcheck). The code is the stop code; a .dmp is in C:\Windows\Minidump.'
	'BugCheck|1001'                              = 'Blue screen (bugcheck). The code is the stop code; a .dmp is in C:\Windows\Minidump.'
	'EventLog|6008'                              = 'The previous shutdown was unexpected.'
	'disk|7'                                     = 'Bad block on a disk - back up and check / replace the drive.'
	'disk|51'                                    = 'Disk paging error - drive, cable or controller problem.'
	'disk|153'                                   = 'Disk I/O was retried - early sign of a failing drive or cable.'
	'Ntfs|55'                                    = 'File system corruption - schedule "Check disk C: with repair at next restart".'
	'Ntfs|98'                                    = 'Volume needs a check disk.'
	'volmgr|161'                                 = 'Crash dump could not be written (page file too small / missing).'
	'WHEA-Logger|1'                              = 'Hardware error reported by the CPU / chipset / memory.'
	'Microsoft-Windows-WHEA-Logger|18'           = 'Fatal hardware error (CPU / memory / bus).'
	'Microsoft-Windows-WHEA-Logger|19'           = 'Corrected hardware error - watch for repeats.'
	'Service Control Manager|7000'               = 'A service failed to start.'
	'Service Control Manager|7001'               = 'A service did not start because a service it depends on failed.'
	'Service Control Manager|7009'               = 'A service took too long to start (timeout).'
	'Service Control Manager|7023'               = 'A service stopped with an error.'
	'Service Control Manager|7031'               = 'A service crashed and was restarted.'
	'Service Control Manager|7034'               = 'A service crashed unexpectedly.'
	'Application Error|1000'                     = 'An application crashed - the faulting program and module are in the description.'
	'Application Hang|1002'                      = 'An application stopped responding and was closed.'
	'.NET Runtime|1026'                          = '.NET application crashed with an unhandled exception.'
	'Microsoft-Windows-DistributedCOM|10016'     = 'DCOM permission warning - harmless, can be ignored.'
	'Schannel|36887'                             = 'TLS alert received from a remote server (protocol / cipher mismatch).'
	'Schannel|36874'                             = 'TLS connection refused - no common protocol / cipher.'
	'Microsoft-Windows-Time-Service|129'         = 'Time sync source not reachable.'
	'Tcpip|4199'                                 = 'IP address conflict on the network.'
	'Microsoft-Windows-WindowsUpdateClient|20'   = 'A Windows update failed to install - the code is in the description.'
}

function Write_EventReport
{
	param ([int]$Days = 7)
	$since = (Get-Date).AddDays(-$Days)
	Set_Status "Reading the event logs (last $Days days)..."
	[System.Windows.Forms.Application]::DoEvents()
	$events = @(Get-WinEvent -FilterHashtable @{ LogName = @('System', 'Application'); Level = @(1, 2); StartTime = $since } -MaxEvents 3000 -ErrorAction SilentlyContinue)
	if ($events.Count -eq 0) { Write_Term "No critical or error events in the System and Application logs in the last $Days days." 'ok'; return $true }

	$groups = @($events | Group-Object { '{0}|{1}|{2}' -f $_.LogName, $_.ProviderName, $_.Id } | ForEach-Object {
			$last = $_.Group | Sort-Object TimeCreated -Descending | Select-Object -First 1
			$level = ($_.Group | Measure-Object Level -Minimum).Minimum
			[pscustomobject]@{ Log = $last.LogName; Provider = $last.ProviderName; Id = $last.Id; Count = $_.Count; Last = $last.TimeCreated; First = ($_.Group | Sort-Object TimeCreated | Select-Object -First 1).TimeCreated; Level = $level; Event = $last }
		} | Sort-Object @{ Expression = 'Level' }, @{ Expression = 'Count'; Descending = $true })
	$crit = @($events | Where-Object { $_.Level -eq 1 }).Count
	Write_Term ('{0} critical and {1} error events in the last {2} days (System + Application), from {3} sources:' -f $crit, ($events.Count - $crit), $Days, $groups.Count) 'info'
	$shown = 0
	foreach ($g in $groups)
	{
		if ($shown -ge 40) { Write_Term ('    ... {0} more sources not shown (open Event Viewer for the rest).' -f ($groups.Count - $shown)) 'muted' -NoTime; break }
		$shown++
		$lvl = if ($g.Level -eq 1) { 'CRITICAL' } else { 'ERROR' }
		$times = if ($g.Count -gt 1) { '{0}x, {1:MM-dd HH:mm} .. {2:MM-dd HH:mm}' -f $g.Count, $g.First, $g.Last } else { '{0:yyyy-MM-dd HH:mm}' -f $g.Last }
		Write_Term '' 'text' -NoTime
		Write_Term ('[{0}]  {1}  |  Event ID {2}  |  {3} log  |  {4}' -f $lvl, $g.Provider, $g.Id, $g.Log, $times) $(if ($g.Level -eq 1) { 'red' } else { 'yellow' }) -NoTime
		$msg = $null
		try { $msg = $g.Event.Message } catch { }
		if (-not $msg) { $msg = '(no description - the provider''s message file is not installed) Data: ' + (($g.Event.Properties | ForEach-Object { [string]$_.Value } | Select-Object -First 4) -join ', ') }
		$desc = (($msg -split "`r?`n" | Where-Object { $_.Trim() }) | Select-Object -First 2) -join ' '
		$desc = ($desc -replace '\s+', ' ').Trim()
		if ($desc.Length -gt 240) { $desc = $desc.Substring(0, 237) + '...' }
		Write_Term ('    ' + $desc) 'text' -NoTime
		$codes = @([regex]::Matches($msg, '0x[0-9A-Fa-f]{6,8}') | ForEach-Object { $_.Value.ToUpper().Replace('0X', '0x') } | Select-Object -Unique -First 4)
		$codes += @([regex]::Matches($msg, '(?i)(error|code|status|c\Wdigo)\s*(code)?\s*[:=]?\s*(-?\d{2,10})\b') | ForEach-Object { $_.Groups[3].Value } | Select-Object -Unique -First 2)
		if ($codes) { Write_Term ('    Codes: ' + ($codes -join ', ')) 'muted' -NoTime }
		$hint = $script:EventHints['{0}|{1}' -f $g.Provider, $g.Id]
		if ($hint) { Write_Term ('    Hint : ' + $hint) 'blue' -NoTime }
	}
	Write_Term '' 'text' -NoTime
	return $true
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Write_DiskHealth - physical disks (health, SMART prediction, reliability) + free space
# ---------------------------------------------------------------------------------------------------
function Write_DiskHealth
{
	$ok = $true
	$predict = @{ }
	try { foreach ($s in @(Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction Stop)) { $predict[[string]$s.InstanceName] = [bool]$s.PredictFailure } } catch { }
	if (@($predict.Values | Where-Object { $_ }).Count) { Write_Term 'SMART: a drive PREDICTS FAILURE - back up now and replace it.' 'fail'; $ok = $false }
	elseif ($predict.Count) { Write_Term ('SMART: {0} drive(s) report no predicted failure.' -f $predict.Count) 'ok' }

	$disks = $null
	try { $disks = @(Get-PhysicalDisk -ErrorAction Stop) } catch { }
	if ($disks)
	{
		foreach ($d in $disks)
		{
			$kind = if ([string]$d.HealthStatus -eq 'Healthy') { 'ok' } else { $ok = $false; 'fail' }
			$sn = ([string]$d.SerialNumber).Trim().TrimEnd('.')
			if ($sn.Length -gt 14) { $sn = '..' + $sn.Substring($sn.Length - 12) }
			Write_Term ('{0,-26} {1,-14} {2,-4} {3,7:N0} GB   health: {4}   status: {5}' -f ([string]$d.FriendlyName).Trim(), $sn, $d.MediaType, ($d.Size / 1GB), $d.HealthStatus, (@($d.OperationalStatus) -join ',')) $kind -NoTime
			try
			{
				$r = $d | Get-StorageReliabilityCounter -ErrorAction Stop
				$parts = @()
				if ($null -ne $r.Temperature -and $r.Temperature -gt 0) { $parts += "temperature $($r.Temperature) C" }
				if ($null -ne $r.Wear) { $parts += "wear $($r.Wear)%" }
				if ($null -ne $r.ReadErrorsUncorrected) { $parts += "uncorrected read errors $($r.ReadErrorsUncorrected)" }
				if ($null -ne $r.PowerOnHours) { $parts += "power-on hours $($r.PowerOnHours)" }
				if ($parts) { Write_Term ('    ' + ($parts -join '   |   ')) $(if ($r.ReadErrorsUncorrected -gt 0 -or $r.Wear -ge 90) { 'warn' } else { 'muted' }) -NoTime }
			}
			catch { }
		}
	}
	else
	{
		foreach ($d in @(Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue))
		{
			Write_Term ('{0,-34} {1,7:N0} GB   status: {2}' -f $d.Model, ($d.Size / 1GB), $d.Status) $(if ($d.Status -eq 'OK') { 'ok' } else { 'fail' }) -NoTime
		}
	}
	foreach ($v in @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction SilentlyContinue))
	{
		if (-not $v.Size) { continue }
		$pct = 100.0 * $v.FreeSpace / $v.Size
		Write_Term ('Drive {0}  {1,7:N1} GB free of {2,7:N1} GB  ({3:N0}% free)' -f $v.DeviceID, ($v.FreeSpace / 1GB), ($v.Size / 1GB), $pct) $(if ($pct -lt 10) { 'warn' } else { 'text' }) -NoTime
	}
	return $ok
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Reset_WindowsUpdate - stop services, rename SoftwareDistribution + catroot2, start again
# ---------------------------------------------------------------------------------------------------
function Reset_WindowsUpdate
{
	$services = @('wuauserv', 'cryptSvc', 'bits', 'msiserver')
	foreach ($s in $services) { try { Stop-Service -Name $s -Force -ErrorAction Stop; Write_Term "Stopped $s" 'muted' -NoTime } catch { Write_Term "Could not stop ${s}: $($_.Exception.Message)" 'warn' -NoTime } }
	$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
	$ok = $true
	foreach ($dir in @((Join-Path $env:SystemRoot 'SoftwareDistribution'), (Join-Path $env:SystemRoot 'System32\catroot2')))
	{
		if (-not (Test-Path -LiteralPath $dir)) { continue }
		try { Rename-Item -LiteralPath $dir -NewName ((Split-Path $dir -Leaf) + '.old_' + $stamp) -ErrorAction Stop; Write_Term "Renamed $dir -> $(Split-Path $dir -Leaf).old_$stamp" 'ok' -NoTime }
		catch { $ok = $false; Write_Term "Could not rename ${dir}: $($_.Exception.Message)" 'fail' -NoTime }
	}
	foreach ($s in $services) { try { Start-Service -Name $s -ErrorAction Stop; Write_Term "Started $s" 'muted' -NoTime } catch { if ($s -ne 'msiserver') { Write_Term "Could not start ${s}: $($_.Exception.Message)" 'warn' -NoTime } } }
	if ($ok) { Write_Term 'Windows Update components reset. Run "DISM - repair" again if it failed before.' 'ok' }
	return $ok
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Invoke_RepairStep - one Repair Windows operation (Start_OpsRun calls it)
# ---------------------------------------------------------------------------------------------------
function Invoke_RepairStep
{
	param ($Step)
	$dism = Get_SystemToolPath 'dism.exe'
	$win8 = ([Environment]::OSVersion.Version -ge [version]'6.2')
	switch ($Step.Function)
	{
		'DismCheck' { if (-not $win8) { Write_Term 'DISM health checks need Windows 8 / Server 2012 or newer - skipped.' 'muted'; return $true }; return (Write_DismResult 'DISM health check' (Invoke_LiveProcess $dism '/Online /Cleanup-Image /CheckHealth')) }
		'DismScan' { if (-not $win8) { Write_Term 'DISM scan needs Windows 8 / Server 2012 or newer - skipped.' 'muted'; return $true }; return (Write_DismResult 'DISM full scan' (Invoke_LiveProcess $dism '/Online /Cleanup-Image /ScanHealth')) }
		'DismRestore' { if (-not $win8) { Write_Term 'DISM repair needs Windows 8 / Server 2012 or newer - skipped.' 'muted'; return $true }; return (Write_DismResult 'DISM repair' (Invoke_LiveProcess $dism '/Online /Cleanup-Image /RestoreHealth /NoRestart')) }
		'DismCleanup' { if (-not $win8) { Write_Term 'Component cleanup needs Windows 8 / Server 2012 or newer - skipped.' 'muted'; return $true }; return (Write_DismResult 'DISM component cleanup' (Invoke_LiveProcess $dism '/Online /Cleanup-Image /StartComponentCleanup /NoRestart')) }
		'Sfc'
		{
			$r = Invoke_LiveProcess (Get_SystemToolPath 'sfc.exe') '/SCANNOW' -Encoding unicode
			$dur = '{0:mm\:ss}' -f $r.Duration
			# English and Spanish wording (the exit code says nothing useful) - same rules as the old worker
			if ($r.Text -match 'did not find any integrity violations|no encontr.{0,40}infracci') { Write_Term "SFC - no damaged system files ($dur)" 'ok'; return $true }
			if ($r.Text -match 'successfully repaired|repar.{0,20}correctamente') { Write_Term "SFC - damaged system files found and REPAIRED ($dur)" 'ok'; return $true }
			if ($r.Text -match 'unable to fix|no pudo reparar|no se pudieron reparar') { Write_Term "SFC - damaged files it could NOT repair: run DISM repair, then SFC again (details: C:\Windows\Logs\CBS\CBS.log) ($dur)" 'fail'; return $false }
			if ($r.Text -match 'pending|pendiente') { Write_Term "SFC - a system repair is pending: restart, then run it again ($dur)" 'warn'; return $false }
			if ($r.Text -match 'could not perform|no pudo realizar') { Write_Term "SFC - could not run the scan ($dur)" 'fail'; return $false }
			if ($r.Exit -eq 0) { Write_Term "SFC - finished ($dur)" 'ok'; return $true }
			Write_Term ("SFC - finished with exit code {0} ($dur)" -f $r.Exit) 'warn'; return $false
		}
		'ChkScan'
		{
			$chk = Get_SystemToolPath 'chkdsk.exe'
			$all = $true
			$vols = @(Get-CimInstance Win32_Volume -Filter "DriveType=3 AND FileSystem='NTFS'" -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter } | Sort-Object DriveLetter)
			foreach ($v in $vols)
			{
				$d = $v.DriveLetter
				Write_Term "Check disk $d" 'info'
				$r = Invoke_LiveProcess $chk $(if ($win8) { "$d /scan" } else { $d })
				switch ($r.Exit)
				{
					0 { Write_Term "$d check disk - no errors" 'ok' }
					1 { Write_Term "$d check disk - errors found and fixed" 'ok' }
					2 { Write_Term "$d check disk - cleanup done" 'ok' }
					3 { Write_Term "$d check disk - problems it can only fix at a restart: run 'Check disk C: with repair at next restart'" 'fail'; $all = $false }
					default { Write_Term ("$d check disk - exit code {0}" -f $r.Exit) 'warn'; $all = $false }
				}
			}
			if (-not $vols) { Write_Term 'No fixed NTFS drive found.' 'muted' }
			return $all
		}
		'DiskHealth' { return (Write_DiskHealth) }
		'Events'
		{
			$days = Read_Prompt 'Event Viewer - how many days back?' '7' '^\d{1,3}$' '1 to 365 days.'
			if ($null -eq $days) { return $false }
			return (Write_EventReport -Days ([Math]::Max(1, [Math]::Min(365, [int]$days))))
		}
		'WuReset' { return (Reset_WindowsUpdate) }
		'ChkdskReboot'
		{
			$chk = Get_SystemToolPath 'chkdsk.exe'
			# chkdsk asks "(Y/N)" in the Windows language: answer No once to read the prompt, then answer
			# with its own "yes" letter (the old tool always sent Y, which fails on Spanish Windows)
			$probe = Invoke_LiveProcess $chk 'C: /f /r' -StdIn 'N' -Quiet
			$m = [regex]::Match($probe.Text, '\((\w)/(\w)\)\s*\??\s*$', 'Multiline')
			if (-not $m.Success) { $m = [regex]::Match($probe.Text, '\((\w)/(\w)\)') }
			$yes = if ($m.Success) { $m.Groups[1].Value } else { 'Y' }
			$null = Invoke_LiveProcess $chk 'C: /f /r' -StdIn $yes
			$boot = [string](@((Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name BootExecute -ErrorAction SilentlyContinue).BootExecute) -join ' ')
			if ($boot -match '(?i)autochk.*/r.*C:') { Write_Term 'C: check disk with repair is scheduled for the next restart (can take hours on a large drive).' 'ok'; return $true }
			Write_Term 'C: check disk could NOT be scheduled - see the output above.' 'fail'
			return $false
		}
	}
	return $true
}

function Initialize_RepairContext { }

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Show_Repair_Info - OS, uptime, pending restart, C: space
# ---------------------------------------------------------------------------------------------------
function Show_Repair_Info
{
	Write_Term 'Repair Windows - this machine' 'step'
	$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
	if ($os)
	{
		$build = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue)
		$ubr = if ($build.UBR) { ".$($build.UBR)" } else { '' }
		Write_Term ('Windows  : {0}  (build {1}{2}, {3})' -f $os.Caption, $os.BuildNumber, $ubr, $os.OSArchitecture) 'text' -NoTime
		$up = (Get-Date) - $os.LastBootUpTime
		Write_Term ('Uptime   : {0}d {1}h {2}m  (last boot {3:yyyy-MM-dd HH:mm})' -f $up.Days, $up.Hours, $up.Minutes, $os.LastBootUpTime) 'text' -NoTime
	}
	$pending = @()
	if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pending += 'component servicing' }
	if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pending += 'Windows Update' }
	if ((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations) { $pending += 'file renames' }
	Write_Term ('Restart  : {0}' -f $(if ($pending) { 'PENDING (' + ($pending -join ', ') + ') - restart before DISM / SFC for best results' } else { 'not pending' })) $(if ($pending) { 'warn' } else { 'text' }) -NoTime
	$c = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
	if ($c) { Write_Term ('Drive C: : {0:N1} GB free of {1:N1} GB' -f ($c.FreeSpace / 1GB), ($c.Size / 1GB)) $(if ($c.FreeSpace -lt 10GB) { 'warn' } else { 'text' }) -NoTime }
	Write_Term 'Recommended order is already set: DISM check, scan, repair, cleanup -> SFC -> check disk -> disk health -> events.' 'muted' -NoTime
}

# ===================================================================================================
#                                     SECTION: Scheduler tool
# ---------------------------------------------------------------------------------------------------
#   Installs the SMS-scheduled DB maintenance task for THIS machine (same files as the old tool):
#     Server: Office\SERVER_DB_MAINTENANCE.SQI + Office\XF<store>901\Add_ServerDBMaintenance_to_RUN_TAB.sqi
#     Lane  : <lane>\Office\LANE_DB_MAINTENANCE.SQI + \\<server>\storeman\office\XF<store><lane>\
#             Add_LaneDBMaintenance_to_RUN_TAB.sqi (the server outbox for this lane, as the old tool)
#   The scripts are the old tool's $ScheduleServerScript / $LaneSQLScript (SMS macro form).
# ===================================================================================================

$script:SqlScheduleServer = @'
/* Set a long timeout so the entire script runs */
@WIZRPL(DBASE_TIMEOUT=E);

/* Set memory configuration */
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', 8192;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000)
DECLARE cmds CURSOR FOR
SELECT 'drop table [' + name + ']'
FROM sys.tables
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%'{{STORE_TABLES}}) AND DATEDIFF(DAY, create_date, GETDATE()) > 30
OPEN cmds
WHILE 1 = 1
BEGIN
FETCH cmds INTO @cmd
IF @@fetch_status != 0 BREAK
EXEC(@cmd)
END
CLOSE cmds;
DEALLOCATE cmds;

/* Shrink database and log files */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'
DBCC SHRINKFILE ({{DB}})
DBCC SHRINKFILE ({{DB}}_Log)
ALTER DATABASE {{DB}} SET RECOVERY SIMPLE

/* Clear the long database timeout */
@WIZCLR(DBASE_TIMEOUT);
'@

$script:SqlScheduleLane = @'
/* Set a long timeout so the entire script runs */
@WIZRPL(DBASE_TIMEOUT=E);

/* Set memory configuration */
EXEC sp_configure 'show advanced options', 1;
RECONFIGURE;
EXEC sp_configure 'max server memory (MB)', 1024;
RECONFIGURE;
EXEC sp_configure 'show advanced options', 0;
RECONFIGURE;

/* Create "{{SQLUSER}}" user in the database */
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = '{{SQLUSER}}')
    CREATE LOGIN [{{SQLUSER}}] WITH PASSWORD = '{{SQLPASS}}', CHECK_POLICY = OFF;
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '{{SQLUSER}}')
    CREATE USER [{{SQLUSER}}] FOR LOGIN [{{SQLUSER}}];
IF IS_ROLEMEMBER('db_owner', '{{SQLUSER}}') = 0
    EXEC sp_addrolemember 'db_owner', '{{SQLUSER}}';

/* Truncate unnecessary tables */
IF OBJECT_ID('COST_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('COST_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE COST_REV;
IF OBJECT_ID('POS_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('POS_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE POS_REV;
IF OBJECT_ID('OBJ_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('OBJ_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE OBJ_REV;
IF OBJECT_ID('PRICE_REV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_REV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE PRICE_REV;
IF OBJECT_ID('REV_HDR', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('REV_HDR', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE REV_HDR;
IF OBJECT_ID('SAL_REG_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_REG_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_REG_SAV;
IF OBJECT_ID('SAL_HDR_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_HDR_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_HDR_SAV;
IF OBJECT_ID('SAL_TTL_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_TTL_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_TTL_SAV;
IF OBJECT_ID('SAL_DET_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('SAL_DET_SAV', 'OBJECT', 'ALTER') = 1 TRUNCATE TABLE SAL_DET_SAV;

/* Truncate PRICE_EVENT table for records older than 7 days */
IF OBJECT_ID('PRICE_EVENT','U') IS NOT NULL AND HAS_PERMS_BY_NAME('PRICE_EVENT','OBJECT','DELETE') = 1 DELETE FROM PRICE_EVENT WHERE F254 < DATEADD(DAY,-7,GETDATE());

/* Drop specific tables older than 30 days */
DECLARE @cmd varchar(4000)
DECLARE cmds CURSOR FOR
SELECT 'drop table [' + name + ']'
FROM sys.tables
WHERE (name LIKE 'TMP_%' OR name LIKE 'MSVHOST%' OR name LIKE 'MMPHOST%'{{STORE_TABLES}}) AND DATEDIFF(DAY, create_date, GETDATE()) > 30
OPEN cmds
WHILE 1 = 1
BEGIN
FETCH cmds INTO @cmd
IF @@fetch_status != 0 BREAK
EXEC(@cmd)
END
CLOSE cmds;
DEALLOCATE cmds;

/* Cleaning HEADER_SAV */
IF OBJECT_ID('HEADER_SAV', 'U') IS NOT NULL AND HAS_PERMS_BY_NAME('HEADER_SAV', 'OBJECT', 'DELETE') = 1
    DELETE FROM HEADER_SAV
    WHERE (F903 = 'SVHOST' OR F903 = 'MPHOST' OR F903 = CONCAT('M', '{{STORE}}', '901'))
    AND (DATEDIFF(DAY, F907, GETDATE()) > 30 OR DATEDIFF(DAY, F909, GETDATE()) > 30);

/* Delete bad SMS items */
@dbEXEC(DELETE FROM OBJ_TAB WHERE F01='0020000000000')
@dbEXEC(DELETE FROM OBJ_TAB WHERE F01 LIKE '% %')
@dbEXEC(DELETE FROM OBJ_TAB WHERE LEN(F01)<>13)
@dbEXEC(DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000')
@dbEXEC(DELETE FROM OBJ_TAB WHERE SUBSTRING(F01,1,3) = '002' AND ISNUMERIC(SUBSTRING(F01,4,5))=0 AND SUBSTRING(F01,9,5) = '00000')
@dbEXEC(DELETE FROM POS_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB))
@dbEXEC(DELETE FROM PRICE_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB))
@dbEXEC(DELETE FROM COST_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB))
@dbEXEC(DELETE FROM SCL_TAB WHERE F01 NOT IN (SELECT F01 FROM OBJ_TAB))
@dbEXEC(DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) <> '002')
@dbEXEC(DELETE FROM SCL_TAB WHERE SUBSTRING(F01,1,3) = '002' AND SUBSTRING(F01,9,5) > '00000')
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F267 = SCL_TXT.F267 FROM SCL_TAB SCL JOIN SCL_TXT_TAB SCL_TXT ON (SCL.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')))
@dbEXEC(UPDATE SCL_TAB SET SCL_TAB.F268 = SCL_NUT.F268 FROM SCL_TAB SCL JOIN SCL_NUT_TAB SCL_NUT ON (SCL.F01=CONCAT('002',FORMAT(SCL_NUT.F268,'00000'),'00000')))
@dbEXEC(DELETE FROM SCL_TXT_TAB WHERE F267 NOT IN (SELECT F267 FROM SCL_TAB))
@dbEXEC(DELETE FROM SCL_NUT_TAB WHERE F268 NOT IN (SELECT F268 FROM SCL_TAB))
@dbEXEC(UPDATE SCL_TAB SET F267 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F267,'00000'),'00000') FROM SCL_TXT_TAB))
@dbEXEC(UPDATE SCL_TAB SET F268 = NULL WHERE F01 NOT IN (SELECT CONCAT('002',FORMAT(F268,'00000'),'00000') FROM SCL_NUT_TAB))
@dbEXEC(UPDATE SCL_TXT_TAB SET SCL_TXT_TAB.F04 = POS.F04 FROM SCL_TXT_TAB SCL_TXT JOIN POS_TAB POS ON (POS.F01=CONCAT('002',FORMAT(SCL_TXT.F267,'00000'),'00000')) WHERE ISNUMERIC(SCL_TXT.F04)=0)
@dbEXEC(UPDATE SCL_TAB SET F256 = REPLACE(REPLACE(REPLACE(F256, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))
@dbEXEC(UPDATE SCL_TAB SET F1952 = REPLACE(REPLACE(REPLACE(F1952, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))
@dbEXEC(UPDATE SCL_TAB SET F2581 = REPLACE(REPLACE(REPLACE(F2581, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))
@dbEXEC(UPDATE SCL_TAB SET F2582 = REPLACE(REPLACE(REPLACE(F2582, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))
@dbEXEC(UPDATE SCL_TXT_TAB SET F297 = REPLACE(REPLACE(REPLACE(F297, CHAR(13),' '), CHAR(10),' '), CHAR(9),' '))

/* Delete batches older than 7 days from the lane */
DELETE FROM Header_bat WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_dct WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_old WHERE F909 < DATEADD(day, -7, GETDATE());
DELETE FROM Header_sav WHERE F909 < DATEADD(day, -7, GETDATE());

/* Rebuild indexes and update database statistics */
EXEC sp_MSforeachtable 'ALTER INDEX ALL ON ? REBUILD'
EXEC sp_MSforeachtable 'UPDATE STATISTICS ? WITH FULLSCAN'

/* Shrink the main database file */
DBCC SHRINKFILE ({{DB}})

/* Shrink the database log file */
DBCC SHRINKFILE ({{DB}}_Log)

/* Restrict the indefinite log file growth */
ALTER DATABASE {{DB}} SET RECOVERY SIMPLE

/* Clear the long database timeout */
@WIZCLR(DBASE_TIMEOUT);
'@

$script:SchedulerSteps = @(
	@{ Name = 'Install DB maintenance task (this machine)'; Function = 'InstallTask'; Default = $true
		Info = 'Server: Office\SERVER_DB_MAINTENANCE.SQI + RUN_TAB task via XF<store>901. Lane: Office\LANE_DB_MAINTENANCE.SQI + RUN_TAB task dropped in the server outbox for this lane. Asks how many days between runs (7 by default).' },
	@{ Name = 'Show the scheduled maintenance task'; Function = 'ShowTask'; Default = $true
		Info = 'Lists the DB maintenance rows of RUN_TAB in this database (task, host, every N days, last / next run).' }
)

# Store server machine name as this lane's database knows it (TER_TAB 901 row, path '@SERVER')
function Get_StoreServerName
{
	$cn = New-Object System.Data.SqlClient.SqlConnection $script:Db.ConnectionString
	try
	{
		$cn.Open()
		$cmd = $cn.CreateCommand()
		$cmd.CommandText = "IF OBJECT_ID('TER_TAB','U') IS NOT NULL SELECT TOP 1 LTRIM(RTRIM(CONVERT(varchar(200), F1125))) FROM TER_TAB WHERE LTRIM(RTRIM(CONVERT(varchar(10), F1057))) = '901' ORDER BY CASE WHEN LTRIM(RTRIM(CONVERT(varchar(10), F1056))) = @s THEN 0 ELSE 1 END"
		$null = $cmd.Parameters.AddWithValue('@s', [string]$script:Db.Store)
		$v = [string]$cmd.ExecuteScalar()
		$v = ($v -replace '^@', '' -replace '^\\\\', '').Trim()
		if ($v -match '[\\/]') { $v = ($v -split '[\\/]')[0] }
		return $v
	}
	catch { return $null }
	finally { $cn.Dispose() }
}

function Invoke_SchedulerStep
{
	param ($Step)
	$d = $script:Db
	if (-not ($d -and $d.Ok)) { Write_Term 'Not connected to the database.' 'fail'; return $false }
	if ($Step.Function -eq 'ShowTask')
	{
		$cn = New-Object System.Data.SqlClient.SqlConnection $d.ConnectionString
		try
		{
			$cn.Open()
			$cmd = $cn.CreateCommand()
			$cmd.CommandText = "IF OBJECT_ID('RUN_TAB','U') IS NOT NULL SELECT F1102, F1000, F1103, F1109, F1114, F1105, F1107 FROM RUN_TAB WHERE F1103 LIKE 'sqi=%DB_MAINTENANCE%' ORDER BY F1000"
			$rd = $cmd.ExecuteReader()
			$n = 0
			while ($rd.Read())
			{
				$n++
				Write_Term ('Task {0}  host {1}  "{2}"  {3}  every {4} day(s)  last {5}  next {6}' -f $rd.GetValue(0), $rd.GetValue(1), $rd.GetValue(3), $rd.GetValue(2), $rd.GetValue(4), $rd.GetValue(5), $rd.GetValue(6)) 'ok' -NoTime
			}
			$rd.Close()
			if (-not $n) { Write_Term 'No DB maintenance task in RUN_TAB yet (a new task appears after SMS processes the scheduler file).' 'warn' }
			return $true
		}
		catch { Write_Term ('Could not read RUN_TAB: {0}' -f $_.Exception.Message) 'fail'; return $false }
		finally { $cn.Dispose() }
	}

	# ---- InstallTask
	if (-not $d.Store) { Write_Term 'Store number unknown (Startup.ini STORE=) - cannot build the task.' 'fail'; return $false }
	if (-not $d.Storeman) { Write_Term 'Storeman folder not found on this machine.' 'fail'; return $false }
	$days = Read_Prompt 'How many days between each run?' '7' '^\d{1,3}$' 'Minimum 1 (7 = weekly).'
	if ($null -eq $days) { return $false }
	$RepeatDays = [Math]::Max(1, [int]$days)

	if ($d.Role -eq 'Server')
	{
		$terminal = '901'
		$sqlText = Expand_Sql $script:SqlScheduleServer
		$destScript = Join-Path $d.Storeman 'Office\SERVER_DB_MAINTENANCE.SQI'
		$xf = Join-Path $d.Storeman ('Office\XF{0}{1}' -f $d.Store, $terminal)
		$macroName = 'Add_ServerDBMaintenance_to_RUN_TAB.sqi'
		$command = 'sqi=SERVER_DB_MAINTENANCE'
		$taskName = 'Server DB Maintenance'
	}
	elseif ($d.Role -eq 'Lane')
	{
		$terminal = [string]$d.Terminal
		if ($terminal -notmatch '^\d{3}$') { Write_Term 'Lane number unknown (Startup.ini TER=).' 'fail'; return $false }
		$guess = Get_StoreServerName
		$server = Read_Prompt 'Store server machine name:' ([string]$guess) '^[A-Za-z0-9._-]{1,63}$' 'The server that sends files to this lane (from TER_TAB terminal 901).'
		if (-not $server) { return $false }
		$sqlText = Expand_Sql $script:SqlScheduleLane
		$destScript = Join-Path $d.Storeman 'Office\LANE_DB_MAINTENANCE.SQI'
		$xf = '\\{0}\storeman\office\XF{1}{2}' -f $server, $d.Store, $terminal
		$macroName = 'Add_LaneDBMaintenance_to_RUN_TAB.sqi'
		$command = 'sqi=LANE_DB_MAINTENANCE'
		$taskName = 'Lane DB Maintenance'
	}
	else { Write_Term 'Machine role unknown - cannot decide between the server and lane task.' 'fail'; return $false }

	if (-not (Test-Path -LiteralPath $xf)) { Write_Term "XF folder not found: $xf - the scheduler file was not dropped." 'fail'; return $false }
	$LastRanDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
	$NextRunDate = (Get-Date).AddDays($RepeatDays).ToString('yyyy-MM-dd HH:mm:ss.fff')
	$macro = @"
 /* First delete the scheduled maintenance if it exists */
 DELETE FROM RUN_TAB WHERE F1103 = '$command' AND F1000 = '$terminal';

 /* Insert the scheduled weekly maintenance */
 INSERT INTO RUN_TAB (F1102, F1000, F1103, F1104, F1105, F1107, F1108, F1109, F1111, F1114, F1115, F1117)
 VALUES (750, '$terminal', '$command', '$terminal', '$LastRanDate', '$NextRunDate', 1, '$taskName', 1, $RepeatDays, 0, 0);

 /* Activate the new task right away */
 @EXEC(SQL=ACTIVATE_ACCEPT_SYS);
"@
	try { Write_Sms_File -Path $destScript -Content $sqlText; Write_Term "Wrote the maintenance script: $destScript" 'ok' }
	catch { Write_Term "Could not write ${destScript}: $($_.Exception.Message)" 'fail'; return $false }
	$macroPath = Join-Path $xf $macroName
	try { Write_Sms_File -Path $macroPath -Content $macro; Write_Term "Scheduler file dropped: $macroPath" 'ok' }
	catch { Write_Term "Could not write ${macroPath}: $($_.Exception.Message)" 'fail'; return $false }
	Write_Term ("'{0}' will run every {1} day(s) once SMS processes the scheduler file (check with 'Show the scheduled maintenance task')." -f $taskName, $RepeatDays) 'info'
	return $true
}

function Show_Scheduler_Info
{
	Write_Term 'Scheduler - DB maintenance task' 'step'
	$d = $script:Db
	if (-not ($d -and $d.Ok)) { Write_Term 'Not connected to the database.' 'warn' -NoTime; return }
	$kind = if ($d.Role -eq 'Lane') { "Lane $($d.Terminal): LANE_DB_MAINTENANCE.SQI, RUN_TAB task via the server outbox" } else { 'Store server: SERVER_DB_MAINTENANCE.SQI, RUN_TAB task via XF' + $d.Store + '901' }
	Write_Term ('This machine : {0}' -f $kind) 'text' -NoTime
	$null = Invoke_SchedulerStep @{ Function = 'ShowTask' }
}

# ===================================================================================================
#                                       SECTION: Export tool
# ---------------------------------------------------------------------------------------------------
#   Store-server tool: the old tool's VNC + remote machine info exports (it ran them silently at
#   every start; here they run when you ask, with their log in the terminal).
# ===================================================================================================

$script:ExportSteps = @(
	@{ Name = 'Export VNC files for all nodes'; Function = 'Export_VNC_Files_For_All_Nodes'; Default = $true
		Info = 'UltraVNC .vnc connection files for every lane, scale and backoffice under Desktop\Lanes, Desktop\Scales, Desktop\Backoffices (reads each node''s VNC password when it can).' },
	@{ Name = 'Get all VNC passwords'; Function = 'Get_All_VNC_Passwords'; Default = $false
		Info = 'Reads the UltraVNC password of every lane, backoffice and scale in parallel and reports which use the default and which a custom password.' },
	@{ Name = 'Remote machine info (hardware export)'; Function = 'Get_Remote_Machine_Info'; Default = $false
		Info = 'Hardware / software info (WMI, CIM, registry) from the nodes you pick, written as info files under Desktop\Lanes, Scales and BackOffices.' }
)

function ConvertTo_Hashtable
{
	param ($Map)
	$h = @{ }
	if ($Map) { foreach ($k in @($Map.Keys)) { $h[$k] = $Map[$k] } }
	$h
}

function Invoke_ExportStep
{
	param ($Step)
	if (-not (Confirm_RemoteNodes)) { Write_Term 'No nodes known for this store - skipped.' 'fail'; return $false }
	$fr = $script:FunctionResults
	$lanes = ConvertTo_Hashtable $fr['LaneNumToMachineName']
	$scales = ConvertTo_Hashtable $fr['ScaleCodeToIPInfo']
	$bo = ConvertTo_Hashtable $fr['BackofficeNumToMachineName']
	switch ($Step.Function)
	{
		'Export_VNC_Files_For_All_Nodes'
		{
			$pw = if ($fr['AllVNCPasswords']) { ConvertTo_Hashtable $fr['AllVNCPasswords'] } else { $null }
			if ($pw) { $null = Export_VNC_Files_For_All_Nodes -LaneNumToMachineName $lanes -ScaleCodeToIPInfo $scales -BackofficeNumToMachineName $bo -AllVNCPasswords $pw }
			else { $null = Export_VNC_Files_For_All_Nodes -LaneNumToMachineName $lanes -ScaleCodeToIPInfo $scales -BackofficeNumToMachineName $bo }
		}
		'Get_All_VNC_Passwords'
		{
			$pw = Get_All_VNC_Passwords -LaneNumToMachineName $lanes -ScaleCodeToIPInfo $scales -BackofficeNumToMachineName $bo
			if ($pw)
			{
				$def = @($pw.GetEnumerator() | Where-Object { $_.Value -eq $script:DefaultVNCPassword }).Count
				$none = @($pw.GetEnumerator() | Where-Object { -not $_.Value }).Count
				Write_Term ('VNC passwords: {0} node(s) - {1} default, {2} custom, {3} not readable.' -f $pw.Count, $def, ($pw.Count - $def - $none), $none) 'ok'
				foreach ($e in ($pw.GetEnumerator() | Sort-Object Name)) { if ($e.Value -and $e.Value -ne $script:DefaultVNCPassword) { Write_Term ('    custom password: {0}' -f $e.Name) 'muted' -NoTime } }
			}
		}
		'Get_Remote_Machine_Info' { $null = Get_Remote_Machine_Info }
	}
	return $true
}

function Show_Export_Info
{
	param ([switch]$Refresh)
	Initialize_RemoteContext
	Write_Term 'Export - store nodes' 'step'
	if (-not (Confirm_RemoteNodes -Refresh:$Refresh)) { Write_Term 'No nodes found (TER_TAB / Office\Load\Ter_Load.sql).' 'warn' -NoTime; return }
	$fr = $script:FunctionResults
	Write_Term ('Lanes {0}   |   Scales {1}   |   Backoffices {2}' -f @($fr['LaneNumToMachineName'].Keys).Count, @(if ($fr['ScaleCodeToIPInfo']) { $fr['ScaleCodeToIPInfo'].Keys }).Count, @(if ($fr['BackofficeNumToMachineName']) { $fr['BackofficeNumToMachineName'].Keys }).Count) 'text' -NoTime
	$desk = [Environment]::GetFolderPath('Desktop')
	Write_Term ('Output       : {0}\Lanes, \Scales, \Backoffices' -f $desk) 'text' -NoTime
}

# ===================================================================================================
#                      Database Maintenance extra: Install FUNCTIONS into SMS (server only)
# ===================================================================================================
function Start_InstallSmsFunctions
{
	if ($script:UI.Prompting -or ($script:Run -and -not $script:Run.Sync.Done)) { return }
	if (-not ($script:Db -and $script:Db.Role -eq 'Server')) { Write_Term 'Install SMS functions is for the store server only.' 'warn'; return }
	if (-not (Show_Dialog 'Install SMS functions' "Write Office\DEPLOY_SYS.sql (with the 'Multiple Functions' menu entry) and Office\DEPLOY_MULTI_FCT.sqm on this server?" 'Install' 'Cancel')) { Write_Term 'Cancelled.' 'warn'; return }
	Initialize_RemoteContext
	$script:UI.Prompting = $true
	Set_RunningUi $true
	try { $null = Install_FUNCTIONS_Into_SMS -StoreNumber $script:FunctionResults['StoreNumber'] -OfficePath $script:OfficePath }
	catch { Write_Term ('Install_FUNCTIONS_Into_SMS failed: {0}' -f $_.Exception.Message) 'fail' }
	finally { $script:UI.Prompting = $false; Set_RunningUi $false }
}

# ===================================================================================================
#                          SECTION: Operation tools (shared check-list engine)
# ---------------------------------------------------------------------------------------------------
#   A tool = toolbar (Run selected + details button) + check list of operations + hint. Each tool:
#     Steps      @{ Name; Function; Default; Info }
#     Role       machine role that may open it ('Server' / 'Lane')
#     Prepare    function run before the operations (paths / context)
#     StepFn     function running one operation, returns $true when it worked
#     InfoFn     details shown in the terminal when the tool is opened
# ===================================================================================================

$script:OpsTools = [ordered]@{
	Scheduler = @{
		Title = 'Scheduler'; Role = 'Known'; Steps = $script:SchedulerSteps
		Prepare = 'Initialize_RemoteContext'; StepFn = 'Invoke_SchedulerStep'; InfoFn = 'Show_Scheduler_Info'; InfoText = 'Current task'
		Description = 'Installs the SMS-scheduled DB maintenance task for this machine (server or lane).'
	}
	Scales = @{
		Title = 'Scales Maintenance'; Role = 'Server'; Steps = $script:ScalesSteps
		Prepare = 'Initialize_ScalesContext'; StepFn = 'Invoke_ScalesStep'; InfoFn = 'Show_Scales_Info'; InfoText = 'Server details'
		Description = 'Tick the operations, then Run selected. The table editor and the duplicate monitor open their own window.'
	}
	Remote = @{
		Title = 'Remote Lane Tools'; Role = 'Server'; Steps = $script:RemoteSteps
		Prepare = 'Initialize_RemoteContext'; StepFn = 'Invoke_RemoteStep'; InfoFn = 'Show_Remote_Info'; InfoText = 'Refresh lanes'
		Description = 'Tick the operations, then Run selected. Each one asks which lanes to use in the lane picker.'
	}
	Export = @{
		Title = 'Export'; Role = 'Server'; Steps = $script:ExportSteps
		Prepare = 'Initialize_RemoteContext'; StepFn = 'Invoke_ExportStep'; InfoFn = 'Show_Export_Info'; InfoText = 'Refresh nodes'
		Description = 'VNC connection files, VNC passwords and remote hardware info for the store nodes (files on the Desktop).'
	}
	Windows = @{
		Title = 'Repair Windows'; Role = 'Any'; Steps = $script:WindowsSteps
		Prepare = 'Initialize_RepairContext'; StepFn = 'Invoke_RepairStep'; InfoFn = 'Show_Repair_Info'; InfoText = 'System info'
		Description = 'Runs in the recommended order (DISM -> SFC -> check disk -> reports). DISM and SFC can take 10-30 minutes each.'
	}
}

# Toolbar + check-list card + gap for one operation tool (the caller adds them to the layout)
function New_OpsToolControls
{
	param ([string]$Key)
	$T = $script:T
	$tool = $script:OpsTools[$Key]

	$bar = New_Panel 'Top' $T.Bg 54
	$bar.Visible = $false
	$bar.Tag = $Key
	$desc = New_Label $tool.Description (New-Object System.Drawing.Font('Segoe UI', 8.75)) $T.Muted
	$desc.Location = New-Object System.Drawing.Point(0, 18)
	$run = New_Button 'Run selected' 130 'accent'
	$run.Tag = $Key
	$run.Add_Click({ Start_OpsRun $this.Tag })
	$info = New_Button $tool.InfoText 120 'ghost'
	$info.Tag = $Key
	$info.Add_Click({
			if ($script:UI.Prompting) { return }
			$t = $script:OpsTools[$this.Tag]
			if ((Get-Command $t.InfoFn).Parameters.ContainsKey('Refresh')) { & $t.InfoFn -Refresh } else { & $t.InfoFn }
		})
	$bar.Controls.AddRange(@($desc, $run, $info))
	$bar.Add_Resize({
			$o = $script:UI.Ops[$this.Tag]
			if ($o) { $o.Run.Location = New-Object System.Drawing.Point(($this.Width - 130), 10); $o.Info.Location = New-Object System.Drawing.Point(($this.Width - 256), 10) }
		})

	$card = New_Panel 'Left' $T.Card 0 360
	$card.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10)
	$card.Visible = $false
	$list = New-Object System.Windows.Forms.CheckedListBox
	$list.Dock = 'Fill'
	$list.BackColor = $T.Card
	$list.ForeColor = $T.Text
	$list.BorderStyle = 'None'
	$list.CheckOnClick = $true
	$list.IntegralHeight = $false
	$list.Font = $T.Font
	$list.Tag = $Key
	$list.Add_SelectedIndexChanged({
			$i = $this.SelectedIndex
			if ($i -ge 0) { $script:UI.Ops[$this.Tag].Hint.Text = $script:OpsTools[$this.Tag].Steps[$i].Info }
		})
	$title = New_Label 'OPERATIONS' (New-Object System.Drawing.Font('Segoe UI Semibold', 8.5)) $T.Muted
	$title.AutoSize = $false
	$title.Dock = 'Top'
	$title.Height = 30
	$hint = New_Label 'Select an operation to see what it does.' (New-Object System.Drawing.Font('Segoe UI', 8.75)) $T.Muted
	$hint.AutoSize = $false
	$hint.Dock = 'Bottom'
	$hint.Height = 84
	$card.Controls.Add($list)
	$card.Controls.Add($hint)
	$card.Controls.Add($title)

	$gap = New_Panel 'Left' $T.Bg 0 14
	$gap.Visible = $false
	@{ Toolbar = $bar; Card = $card; Gap = $gap; List = $list; Hint = $hint; Run = $run; Info = $info }
}

function Fill_OpsList
{
	param ([string]$Key)
	$saved = $script:Settings[$Key]
	$list = $script:UI.Ops[$Key].List
	$list.Items.Clear()
	foreach ($s in $script:OpsTools[$Key].Steps)
	{
		$checked = if ($saved -and $saved.ContainsKey($s.Name)) { $saved[$s.Name] } else { [bool]$s.Default }
		[void]$list.Items.Add($s.Name, $checked)
	}
}

function Save_OpsChecked
{
	param ([string]$Key)
	if (-not $script:UI.Ops -or -not $script:UI.Ops[$Key]) { return }
	$list = $script:UI.Ops[$Key].List
	$steps = $script:OpsTools[$Key].Steps
	$h = @{ }
	for ($i = 0; $i -lt $steps.Count; $i++) { $h[$steps[$i].Name] = $list.GetItemChecked($i) }
	$script:Settings[$Key] = $h
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Start_OpsRun - runs the ticked operations of a tool, in list order
# ---------------------------------------------------------------------------------------------------
function Start_OpsRun
{
	param ([string]$Key)
	if ($script:UI.Prompting -or $script:MgJob) { return }
	$tool = $script:OpsTools[$Key]
	$list = $script:UI.Ops[$Key].List
	$sel = @()
	for ($i = 0; $i -lt $tool.Steps.Count; $i++) { if ($list.GetItemChecked($i)) { $sel += $tool.Steps[$i] } }
	if ($sel.Count -eq 0) { Write_Term 'No operations selected.' 'warn'; return }
	Save_OpsChecked $Key
	Save_Settings
	$names = ($sel | ForEach-Object { '  - ' + $_.Name }) -join "`r`n"
	if (-not (Show_Dialog $tool.Title "Run $($sel.Count) operation(s)?`r`n$names" 'Run' 'Cancel')) { Write_Term 'Cancelled.' 'warn'; return }

	& $tool.Prepare
	$script:UI.Prompting = $true
	Set_RunningUi $true
	$started = [datetime]::Now
	$ok = 0; $bad = 0
	try
	{
		$n = 0
		foreach ($s in $sel)
		{
			$n++
			Set_Status ('Running {0}/{1}: {2}' -f $n, $sel.Count, $s.Name) (100.0 * ($n - 1) / $sel.Count)
			Write_Term ('[{0}/{1}] {2}' -f $n, $sel.Count, $s.Name) 'step'
			try
			{
				if (& $tool.StepFn $s) { $ok++ } else { $bad++ }
			}
			catch
			{
				$bad++
				Write_Term ('{0} failed: {1}' -f $s.Name, $_.Exception.Message) 'fail'
			}
		}
	}
	finally
	{
		$script:UI.Prompting = $false
		Set_RunningUi $false
		$script:UI.Elapsed.Text = '{0:hh\:mm\:ss}' -f ([datetime]::Now - $started)
		Set_Status 'Finished' 100
	}
	$kind = if ($bad) { 'warn' } else { 'ok' }
	Write_Term ('{0} finished: {1} OK, {2} with errors.' -f $tool.Title, $ok, $bad) $kind
}


# ===================================================================================================
#                                      SECTION: GUI helpers
# ===================================================================================================

function Initialize_Theme
{
	$c = { param ($r, $g, $b) [System.Drawing.Color]::FromArgb($r, $g, $b) }
	$script:T = @{
		Bg       = & $c 22 22 26
		Side     = & $c 28 28 33
		Card     = & $c 36 36 42
		CardHi   = & $c 46 46 54
		Border   = & $c 58 58 68
		Text     = & $c 230 232 236
		Muted    = & $c 145 150 162
		Accent   = & $c 56 139 253
		AccentHi = & $c 88 160 255
		Ok       = & $c 63 185 80
		Warn     = & $c 210 153 34
		Fail     = & $c 248 81 73
		Info     = & $c 121 192 255
		Danger   = & $c 200 50 48
		TermBg   = & $c 12 14 18
		TermText = & $c 201 209 217
		Sql      = & $c 106 115 128
		Font     = New-Object System.Drawing.Font('Segoe UI', 9.75)
		FontBold = New-Object System.Drawing.Font('Segoe UI Semibold', 9.75)
		FontH    = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
		FontMono = New-Object System.Drawing.Font('Consolas', 10)
		Brand    = & $c 132 204 22      # logo lime green
		BrandDim = & $c 98 160 255
		Navy     = & $c 44 36 92        # logo navy (title letters 1, 3, 5...)
		White    = & $c 248 248 248     # logo white (title letters 2, 4, 6...)
	}
	$script:T.TitleColors = @($script:T.Navy)   # every letter logo navy (add colours here to alternate)
	# white outline on the navy letters: $script:Cfg.TitleOutline (SMSTK_TITLE_OUTLINE=0/1 overrides it for previews)
	$script:T.TitleOutline = if ($env:SMSTK_TITLE_OUTLINE) { $env:SMSTK_TITLE_OUTLINE -eq '1' } else { [bool]$script:Cfg.TitleOutline }
	Initialize_Branding
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: New_ColorTitle - the app name drawn letter by letter, cycling through -Colors
#   (a Label paints one colour only). Returns a transparent Label sized to the text.
# ---------------------------------------------------------------------------------------------------
function New_ColorTitle
{
	param ([string]$Text, $Font, [object[]]$Colors, [switch]$Outline)
	# letter positions = width of everything before it (keeps the font's own spacing / kerning)
	$fmt = [System.Drawing.StringFormat]::GenericTypographic
	$bmp = New-Object System.Drawing.Bitmap(1, 1)
	$mg = [System.Drawing.Graphics]::FromImage($bmp)
	$mg.TextRenderingHint = 'AntiAlias'
	$xs = @(for ($i = 0; $i -lt $Text.Length; $i++) { if ($i -eq 0) { 0.0 } else { $mg.MeasureString($Text.Substring(0, $i), $Font, (New-Object System.Drawing.PointF(0, 0)), $fmt).Width } })
	$full = $mg.MeasureString($Text, $Font, (New-Object System.Drawing.PointF(0, 0)), $fmt)
	$emPx = $Font.SizeInPoints * $mg.DpiY / 72
	$mg.Dispose(); $bmp.Dispose()
	$l = New-Object System.Windows.Forms.Label
	$l.AutoSize = $false
	$l.BackColor = [System.Drawing.Color]::Transparent
	$l.UseMnemonic = $false
	# extra width: the last italic letter leans past its measured box
	$l.Size = New-Object System.Drawing.Size(([int]$full.Width + [int]($full.Height * 0.3) + 4), ([int]$full.Height + 6))
	$l.Tag = @{ Text = $Text; Font = $Font; Colors = $Colors; Xs = $xs; EmPx = $emPx; Outline = [bool]$Outline }
	$l.Add_Paint({
			param ($s, $e)
			$t = $s.Tag
			$g = $e.Graphics
			$g.SmoothingMode = 'AntiAlias'
			$fmt = [System.Drawing.StringFormat]::GenericTypographic
			for ($i = 0; $i -lt $t.Text.Length; $i++)
			{
				$color = $t.Colors[$i % $t.Colors.Count]
				$path = New-Object System.Drawing.Drawing2D.GraphicsPath
				$path.AddString([string]$t.Text[$i], $t.Font.FontFamily, [int]$t.Font.Style, $t.EmPx, (New-Object System.Drawing.PointF(($t.Xs[$i] + 2), 3)), $fmt)
				# dark letters get a thin white outline so they stay readable on the dark background
				if ($t.Outline -and $color.GetBrightness() -lt 0.35)
				{
					$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(235, 248, 248, 248), [single]([Math]::Max(1.2, $t.EmPx / 22)))
					$pen.LineJoin = 'Round'
					$g.DrawPath($pen, $path)
					$pen.Dispose()
				}
				$brush = New-Object System.Drawing.SolidBrush($color)
				$g.FillPath($brush, $path)
				$brush.Dispose(); $path.Dispose()
			}
		})
	$l
}

# Title font: Segoe UI Black Italic (Windows 8+), else Segoe UI Bold Italic
function New_TitleFont
{
	param ([single]$Size)
	$installed = @([System.Drawing.FontFamily]::Families | ForEach-Object { $_.Name })
	if ($installed -contains 'Segoe UI Black') { return New-Object System.Drawing.Font('Segoe UI Black', $Size, [System.Drawing.FontStyle]::Italic) }
	New-Object System.Drawing.Font('Segoe UI', $Size, ([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Initialize_Branding - decodes the embedded logo: $script:T.Logo (Image), $script:T.AppIcon
# ---------------------------------------------------------------------------------------------------
function Initialize_Branding
{
	try
	{
		$png = [Convert]::FromBase64String(($script:LogoPngB64 -replace '\s', ''))
		$script:T.Logo = [System.Drawing.Image]::FromStream((New-Object System.IO.MemoryStream(, $png)))
	}
	catch { $script:T.Logo = $null }
	try
	{
		$ico = [Convert]::FromBase64String(($script:LogoIcoB64 -replace '\s', ''))
		$script:T.AppIcon = New-Object System.Drawing.Icon((New-Object System.IO.MemoryStream(, $ico)))
	}
	catch { $script:T.AppIcon = $null }
}

# Logo as a transparent PictureBox of the given size (nothing when the logo could not be decoded)
function New_LogoBox
{
	param ([int]$Size, [int]$X, [int]$Y)
	if (-not $script:T.Logo) { return $null }
	$p = New-Object System.Windows.Forms.PictureBox
	$p.Image = $script:T.Logo
	$p.SizeMode = 'Zoom'
	$p.BackColor = [System.Drawing.Color]::Transparent
	$p.SetBounds($X, $Y, $Size, $Size)
	$p
}

# Flat dark button
function New_Button
{
	param ([string]$Text, [int]$Width = 110, [string]$Kind = 'normal')
	$b = New-Object System.Windows.Forms.Button
	$b.Text = $Text
	$b.Width = $Width
	$b.Height = 32
	$b.FlatStyle = 'Flat'
	$b.FlatAppearance.BorderSize = 0
	$b.Cursor = [System.Windows.Forms.Cursors]::Hand
	$b.Font = $script:T.FontBold
	$b.ForeColor = $script:T.Text
	switch ($Kind)
	{
		'accent' { $b.BackColor = $script:T.Accent; $b.FlatAppearance.MouseOverBackColor = $script:T.AccentHi; $b.ForeColor = [System.Drawing.Color]::White }
		'danger' { $b.BackColor = $script:T.Danger; $b.FlatAppearance.MouseOverBackColor = $script:T.Fail; $b.ForeColor = [System.Drawing.Color]::White }
		'ghost' { $b.BackColor = $script:T.Card; $b.FlatAppearance.MouseOverBackColor = $script:T.CardHi; $b.Font = $script:T.Font }
		default { $b.BackColor = $script:T.CardHi; $b.FlatAppearance.MouseOverBackColor = $script:T.Border }
	}
	$b
}

function New_Label
{
	param ([string]$Text, $Font = $script:T.Font, $Color = $script:T.Text)
	$l = New-Object System.Windows.Forms.Label
	$l.UseMnemonic = $false   # show '&' (Alex C.T & Enrique P.F) instead of treating it as a hotkey
	$l.Text = $Text
	$l.Font = $Font
	$l.ForeColor = $Color
	$l.BackColor = [System.Drawing.Color]::Transparent
	$l.AutoSize = $true
	$l
}

function New_Panel
{
	param ([string]$Dock = 'None', $Color = $script:T.Bg, [int]$Height = 0, [int]$Width = 0)
	$p = New-Object System.Windows.Forms.Panel
	$p.Dock = $Dock
	$p.BackColor = $Color
	if ($Height) { $p.Height = $Height }
	if ($Width) { $p.Width = $Width }
	$p
}

# Dark modal dialog. -TypeToConfirm makes the OK button wait for that exact word.
function Show_Dialog
{
	param ([string]$Title, [string]$Message, [string]$OkText = 'OK', [string]$CancelText = 'Cancel', [string]$TypeToConfirm, [switch]$Danger)
	$f = New-Object System.Windows.Forms.Form
	$f.FormBorderStyle = 'None'
	$f.StartPosition = 'CenterParent'
	$f.BackColor = if ($Danger) { $script:T.Danger } else { $script:T.Border }
	$f.Padding = New-Object System.Windows.Forms.Padding(1)
	$f.Size = New-Object System.Drawing.Size(520, 260)
	$f.ShowInTaskbar = $false
	$f.KeyPreview = $true

	$in = New_Panel 'Fill' $script:T.Card
	$in.Padding = New-Object System.Windows.Forms.Padding(20, 16, 20, 16)
	$f.Controls.Add($in)

	$body = New_Label $Message $script:T.Font $script:T.Text
	$body.AutoSize = $false
	$body.Dock = 'Fill'
	$bottom = New_Panel 'Bottom' $script:T.Card 44
	$title = New_Label $Title $script:T.FontH $(if ($Danger) { $script:T.Fail } else { $script:T.Text })
	$title.AutoSize = $false
	$title.Dock = 'Top'
	$title.Height = 34
	$in.Controls.Add($body)
	$in.Controls.Add($bottom)
	$in.Controls.Add($title)

	$ok = New_Button $OkText 120 $(if ($Danger) { 'danger' } else { 'accent' })
	$ok.DialogResult = 'OK'
	$ok.Anchor = 'Right,Bottom'
	$cancel = New_Button $CancelText 100
	$cancel.DialogResult = 'Cancel'
	$cancel.Anchor = 'Right,Bottom'
	$bottom.Controls.Add($ok)
	$bottom.Controls.Add($cancel)
	$ok.Location = New-Object System.Drawing.Point(($bottom.Width - 230), 8)
	$cancel.Location = New-Object System.Drawing.Point(($bottom.Width - 104), 8)
	$f.CancelButton = $cancel

	if ($TypeToConfirm)
	{
		$f.Height = 300
		$box = New-Object System.Windows.Forms.TextBox
		$box.BackColor = $script:T.TermBg
		$box.ForeColor = $script:T.Text
		$box.BorderStyle = 'FixedSingle'
		$box.Font = $script:T.FontMono
		$box.Width = 160
		$box.Location = New-Object System.Drawing.Point(0, 10)
		$box.Tag = $TypeToConfirm
		$bottom.Controls.Add($box)
		$ok.Enabled = $false
		$box.Add_TextChanged({ $this.Parent.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.DialogResult -eq 'OK' } | ForEach-Object { $_.Enabled = ($this.Text -ceq $this.Tag) } })
		$f.Add_Shown({ $this.Controls[0].Controls | ForEach-Object { $_.Controls } | Where-Object { $_ -is [System.Windows.Forms.TextBox] } | ForEach-Object { $_.Focus() } })
	}
	else { $f.AcceptButton = $ok }

	$res = $f.ShowDialog($script:UI.Form)
	$f.Dispose()
	return ($res -eq 'OK')
}

# ===================================================================================================
#                                      SECTION: Splash screen
# ===================================================================================================

function New_Splash
{
	$f = New-Object System.Windows.Forms.Form
	$f.FormBorderStyle = 'None'
	$f.StartPosition = 'CenterScreen'
	$f.Size = New-Object System.Drawing.Size(580, 330)
	$f.BackColor = $script:T.Bg
	$f.Text = $script:App.Name
	$f.TopMost = $true
	if ($script:T.AppIcon) { $f.Icon = $script:T.AppIcon }

	$strip = New_Panel 'Top' $script:T.Accent 4
	$f.Controls.Add($strip)

	# logo + title side by side, centred as one group
	$titleFont = New_TitleFont 34
	$title = New_ColorTitle $script:App.Name $titleFont $script:T.TitleColors -Outline:$script:T.TitleOutline
	$tw = $title.Width
	$logoSize = 84
	$gap = if ($script:T.Logo) { 16 } else { 0 }
	$groupW = $(if ($script:T.Logo) { $logoSize } else { 0 }) + $gap + $tw
	$x0 = [int]((580 - $groupW) / 2)
	$logo = New_LogoBox $logoSize $x0 52
	if ($logo) { $f.Controls.Add($logo) }
	$title.Location = New-Object System.Drawing.Point(($x0 + $groupW - $tw), (52 + [int](($logoSize - $title.Height) / 2)))
	$f.Controls.Add($title)

	$tag = New_Label $script:App.Tagline (New-Object System.Drawing.Font('Segoe UI', 11)) $script:T.Muted
	$tag.AutoSize = $false
	$tag.TextAlign = 'MiddleCenter'
	$tag.SetBounds(0, 146, 580, 24)
	$f.Controls.Add($tag)

	$company = New_Label $script:App.Company (New-Object System.Drawing.Font('Segoe UI Semibold', 10)) $script:T.Text
	$company.AutoSize = $false
	$company.TextAlign = 'MiddleCenter'
	$company.SetBounds(0, 176, 580, 22)
	$f.Controls.Add($company)

	$credits = New_Label ('by  ' + $script:App.Credits) (New-Object System.Drawing.Font('Segoe UI Semibold', 10)) $script:T.AccentHi
	$credits.AutoSize = $false
	$credits.TextAlign = 'MiddleCenter'
	$credits.SetBounds(0, 198, 580, 22)
	$f.Controls.Add($credits)

	$status = New_Label 'Starting...' $script:T.Font $script:T.Muted
	$status.AutoSize = $false
	$status.SetBounds(60, 240, 460, 20)
	$f.Controls.Add($status)

	$track = New_Panel 'None' $script:T.Card
	$track.SetBounds(60, 264, 460, 6)
	$fill = New_Panel 'None' $script:T.Accent
	$fill.SetBounds(0, 0, 0, 6)
	$track.Controls.Add($fill)
	$f.Controls.Add($track)

	$ver = New_Label ('v' + $script:App.Version) $script:T.Font $script:T.Muted
	$ver.Location = New-Object System.Drawing.Point(520, 300)
	$f.Controls.Add($ver)

	@{ Form = $f; Status = $status; Fill = $fill; Track = $track; Value = 0.0 }
}

# Moves the splash bar toward Percent (smooth) and pumps the message loop
function Update_Splash
{
	param ($Splash, [double]$Percent, [string]$Text)
	if (-not $Splash) { return }
	if ($Text) { $Splash.Status.Text = $Text }
	$target = [Math]::Min(100, $Percent)
	do
	{
		$Splash.Value = [Math]::Min($target, $Splash.Value + [Math]::Max(1.5, ($target - $Splash.Value) / 3))
		$Splash.Fill.Width = [int]($Splash.Track.Width * $Splash.Value / 100)
		$Splash.Form.Refresh()
		[System.Windows.Forms.Application]::DoEvents()
		if ($Splash.Value -lt $target) { Start-Sleep -Milliseconds 8 }
	} while ($Splash.Value -lt $target)
}

# ===================================================================================================
#                                       SECTION: Main window
# ===================================================================================================

function New_MainWindow
{
	$T = $script:T
	$f = New-Object System.Windows.Forms.Form
	$f.Text = $script:App.Name
	$f.FormBorderStyle = 'None'
	$f.StartPosition = 'CenterScreen'
	$f.Size = New-Object System.Drawing.Size(1280, 780)
	$f.MinimumSize = New-Object System.Drawing.Size(1000, 620)
	$f.BackColor = $T.Border
	$f.Padding = New-Object System.Windows.Forms.Padding(1)
	$f.Font = $T.Font
	$f.KeyPreview = $true
	if ($T.AppIcon) { $f.Icon = $T.AppIcon }   # taskbar + Alt-Tab
	$script:UI.Form = $f

	$root = New_Panel 'Fill' $T.Bg
	$f.Controls.Add($root)

	# ---------------------------------------------------------------- content (Fill, added first)
	$content = New_Panel 'Fill' $T.Bg
	$content.Padding = New-Object System.Windows.Forms.Padding(18, 14, 18, 10)
	$root.Controls.Add($content)

	# ---------------------------------------------------------------- status bar (Bottom)
	$status = New_Panel 'Bottom' $T.Side 30
	$status.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0)
	$statusText = New_Label 'Ready' $T.Font $T.Muted
	$statusText.AutoSize = $false
	$statusText.Dock = 'Fill'
	$statusText.TextAlign = 'MiddleLeft'
	$barTrack = New_Panel 'Right' $T.Side 0 260
	$barInner = New_Panel 'None' $T.Card
	$barInner.SetBounds(0, 12, 260, 6)
	$barFill = New_Panel 'None' $T.Accent
	$barFill.SetBounds(0, 0, 0, 6)
	$barInner.Controls.Add($barFill)
	$barTrack.Controls.Add($barInner)
	$elapsed = New_Label '' $T.Font $T.Muted
	$elapsed.AutoSize = $false
	$elapsed.Dock = 'Right'
	$elapsed.Width = 90
	$elapsed.TextAlign = 'MiddleRight'
	$status.Controls.Add($statusText)
	$status.Controls.Add($elapsed)
	$status.Controls.Add($barTrack)
	$root.Controls.Add($status)
	$script:UI.StatusText = $statusText
	$script:UI.BarFill = $barFill
	$script:UI.BarInner = $barInner
	$script:UI.Elapsed = $elapsed

	# ---------------------------------------------------------------- sidebar (Left)
	$side = New_Panel 'Left' $T.Side 0 220
	$side.Padding = New-Object System.Windows.Forms.Padding(0, 16, 0, 12)
	# bottom of the sidebar, top to bottom: version / Tecnica Systems / names
	# (Dock=Bottom: the last one added sits lowest)
	foreach ($spec in @(@(('v' + $script:App.Version), $T.Font, $T.Muted), @($script:App.Company, $T.FontBold, $T.Text), @($script:App.Credits, $T.Font, $T.Muted)))
	{
		$l = New_Label $spec[0] $spec[1] $spec[2]
		$l.AutoSize = $false
		$l.Dock = 'Bottom'
		$l.Height = 20
		$l.Padding = New-Object System.Windows.Forms.Padding(18, 0, 0, 0)
		$side.Controls.Add($l)
	}

	$toolsTitle = New_Label '  TOOLS' (New-Object System.Drawing.Font('Segoe UI Semibold', 8.5)) $T.Muted
	$toolsTitle.AutoSize = $false
	$toolsTitle.Dock = 'Top'
	$toolsTitle.Height = 26
	$toolsTitle.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)

	$navButtons = [ordered]@{ }
	$navSpecs = @(@('Db', '   Database Maintenance'), @('MiniGhost', '   Mini Ghost'))
	foreach ($k in $script:OpsTools.Keys) { $navSpecs += , @($k, ('   ' + $script:OpsTools[$k].Title)) }
	foreach ($spec in $navSpecs)
	{
		$nb = New-Object System.Windows.Forms.Button
		$nb.Text = $spec[1]
		$nb.Tag = $spec[0]
		$nb.TextAlign = 'MiddleLeft'
		$nb.Dock = 'Top'
		$nb.Height = 40
		$nb.FlatStyle = 'Flat'
		$nb.FlatAppearance.BorderSize = 0
		$nb.FlatAppearance.MouseOverBackColor = $T.Card
		$nb.BackColor = $T.Side
		$nb.ForeColor = $T.Text
		$nb.Font = $T.FontBold
		$nb.Cursor = [System.Windows.Forms.Cursors]::Hand
		$mark = New_Panel 'Left' $T.Accent 0 4
		$mark.Visible = $false
		$nb.Controls.Add($mark)
		$nb.Add_Click({ Select_Tool $this.Tag })
		$navButtons[$spec[0]] = @{ Button = $nb; Mark = $mark }
	}
	$script:UI.NavDb = $navButtons['Db'].Button
	$script:UI.NavDbMark = $navButtons['Db'].Mark
	$script:UI.NavMg = $navButtons['MiniGhost'].Button
	$script:UI.NavMgMark = $navButtons['MiniGhost'].Mark
	$script:UI.Nav = $navButtons
	$script:UI.NavDb.BackColor = $T.CardHi
	$script:UI.NavDbMark.Visible = $true
	$script:UI.NavMg.Enabled = $false   # enabled by Complete_Detection on a lane
	foreach ($k in $script:OpsTools.Keys) { $navButtons[$k].Button.Enabled = $false }   # by machine role
	$script:UI.Tip = New-Object System.Windows.Forms.ToolTip
	$script:UI.Tip.SetToolTip($script:UI.NavMg, 'Lane only - detecting the database...')

	$soon = New_Label '   More tools coming soon' $T.Font $T.Muted
	$soon.AutoSize = $false
	$soon.Dock = 'Top'
	$soon.Height = 34
	$soon.TextAlign = 'MiddleLeft'

	# Dock=Top stacks in reverse order of adding
	$side.Controls.Add($soon)
	foreach ($k in @($navButtons.Keys)[($navButtons.Count - 1)..0]) { $side.Controls.Add($navButtons[$k].Button) }
	$side.Controls.Add($toolsTitle)
	$root.Controls.Add($side)

	# ---------------------------------------------------------------- title bar (Top, borderless window)
	$head = New_Panel 'Top' $T.Side 42
	$headLogo = New_LogoBox 26 14 8
	$titleX = if ($headLogo) { 48 } else { 16 }
	$headFont = New_TitleFont 13
	$appTitle = New_ColorTitle $script:App.Name $headFont $T.TitleColors -Outline:$T.TitleOutline
	$appTitle.Location = New-Object System.Drawing.Point($titleX, [int]((42 - $appTitle.Height) / 2))
	$titleW = $appTitle.Width
	$sub = New_Label '|  Database Maintenance' $T.Font $T.Muted
	$sub.Location = New-Object System.Drawing.Point(($titleX + $titleW + 6), 13)
	if ($headLogo) { $head.Controls.Add($headLogo) }
	$head.Controls.Add($appTitle)
	$head.Controls.Add($sub)
	$script:UI.Sub = $sub
	# Dock=Right: the last one added sits furthest right -> minimize, maximize, close (X on the far right)
	foreach ($spec in @(@('min', [char]0x2014), @('max', [char]0x25A1), @('close', [char]0x2715)))
	{
		$b = New-Object System.Windows.Forms.Button
		$b.Text = [string]$spec[1]
		$b.Tag = $spec[0]
		$b.Dock = 'Right'
		$b.Width = 46
		$b.FlatStyle = 'Flat'
		$b.FlatAppearance.BorderSize = 0
		$b.ForeColor = $T.Text
		$b.BackColor = $T.Side
		$b.Font = New-Object System.Drawing.Font('Segoe UI', 10)
		$b.FlatAppearance.MouseOverBackColor = if ($spec[0] -eq 'close') { $T.Danger } else { $T.CardHi }
		$b.Add_Click({
				switch ($this.Tag)
				{
					'close' { $script:UI.Form.Close() }
					'min' { $script:UI.Form.WindowState = 'Minimized' }
					'max' { Switch_Maximize }
				}
			})
		$head.Controls.Add($b)
	}
	foreach ($c in @($head, $appTitle, $sub, $headLogo | Where-Object { $_ }))
	{
		$c.Add_MouseDown({ if ($_.Button -eq 'Left') { $script:UI.Drag = @{ Mouse = [System.Windows.Forms.Cursor]::Position; Form = $script:UI.Form.Location } } })
		$c.Add_MouseMove({
				if ($script:UI.Drag -and -not $script:UI.NormalBounds)
				{
					$p = [System.Windows.Forms.Cursor]::Position
					$script:UI.Form.Location = New-Object System.Drawing.Point(($script:UI.Drag.Form.X + $p.X - $script:UI.Drag.Mouse.X), ($script:UI.Drag.Form.Y + $p.Y - $script:UI.Drag.Mouse.Y))
				}
			})
		$c.Add_MouseUp({ $script:UI.Drag = $null })
		$c.Add_DoubleClick({ Switch_Maximize })
	}
	$root.Controls.Add($head)

	# ================================================================ content area
	# body (Fill) = query list (Left) + terminal (Fill); above it the toolbar and the connection card
	$body = New_Panel 'Fill' $T.Bg
	$content.Controls.Add($body)

	# ---- toolbar: mode switch + actions
	$toolbar = New_Panel 'Top' $T.Bg 54
	$toolbar.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 10)
	$modeLabel = New_Label 'MODE' (New-Object System.Drawing.Font('Segoe UI Semibold', 8.5)) $T.Muted
	$modeLabel.Location = New-Object System.Drawing.Point(0, 19)
	$toolbar.Controls.Add($modeLabel)
	$x = 48
	$script:UI.ModeButtons = @{ }
	foreach ($m in @(@('Maintenance', 'Maintenance'), @('Repair', 'Repair'), @('Custom', 'Load .sqi...')))
	{
		$b = New_Button $m[1] 124 'ghost'
		$b.Tag = $m[0]
		$b.Location = New-Object System.Drawing.Point($x, 10)
		$b.Add_Click({ Select_Mode $this.Tag })
		$toolbar.Controls.Add($b)
		$script:UI.ModeButtons[$m[0]] = $b
		$x += 128
	}
	# server only: Install_FUNCTIONS_Into_SMS (shown by Update_MiniGhostNav on the store server)
	$smsFn = New_Button 'Install SMS functions' 170 'ghost'
	$smsFn.Location = New-Object System.Drawing.Point(($x + 16), 10)
	$smsFn.Visible = $false
	$smsFn.Add_Click({ Start_InstallSmsFunctions })
	$toolbar.Controls.Add($smsFn)
	$script:UI.SmsFn = $smsFn
	$run = New_Button 'Run selected' 130 'accent'
	$run.Anchor = 'Top,Right'
	$stop = New_Button 'Stop' 90 'danger'
	$stop.Anchor = 'Top,Right'
	$stop.Enabled = $false
	$showSql = New-Object System.Windows.Forms.CheckBox
	$showSql.Text = 'Show SQL'
	$showSql.Checked = $true
	$showSql.ForeColor = $T.Muted
	$showSql.AutoSize = $true
	$showSql.Anchor = 'Top,Right'
	$toolbar.Controls.Add($run)
	$toolbar.Controls.Add($stop)
	$toolbar.Controls.Add($showSql)
	$toolbar.Add_Resize({
			$w = $this.Width
			$script:UI.Run.Location = New-Object System.Drawing.Point(($w - 130), 10)
			$script:UI.Stop.Location = New-Object System.Drawing.Point(($w - 226), 10)
			$script:UI.ShowSql.Location = New-Object System.Drawing.Point(($w - 320), 16)
		})
	$run.Add_Click({ Start_SelectedRun })
	$stop.Add_Click({ Stop_Run })
	$script:UI.Run = $run
	$script:UI.Stop = $stop
	$script:UI.ShowSql = $showSql
	$script:UI.Toolbar = $toolbar

	# ---- Mini Ghost actions (shown instead of the toolbar + query list)
	$mg = New_Panel 'Top' $T.Bg 112
	$mg.Visible = $false
	$mgRun = New_Button 'Run Mini-Ghost' 230 'accent'
	$mgRun.Height = 40
	$mgRun.Location = New-Object System.Drawing.Point(0, 14)
	$mgRun.Add_Click({ Start_MiniGhost_Run })
	$mgDb = New_Button 'Update Store SQL Database' 230
	$mgDb.Height = 40
	$mgDb.Location = New-Object System.Drawing.Point(242, 14)
	$mgDb.Add_Click({ Start_MiniGhost_DbUpdate })
	$mgInfoBtn = New_Button 'Lane details' 120 'ghost'
	$mgInfoBtn.Height = 40
	$mgInfoBtn.Location = New-Object System.Drawing.Point(484, 14)
	$mgInfoBtn.Add_Click({ if (-not $script:MgJob -and -not $script:UI.Prompting) { Show_MiniGhost_Info } })
	$mgText = New_Label ("Run Mini-Ghost: this lane was restored from another lane's image - asks store #, lane # and IP, then renames it, syncs the INIs and SQL, truncates the sales tables, clears the GT counters, sets the IP and offers a reboot.`r`nUpdate Store SQL Database: only the database was copied from another lane - updates the SQL tables for this lane and truncates. No rename, no INI, no IP.") (New-Object System.Drawing.Font('Segoe UI', 8.75)) $T.Muted
	$mgText.AutoSize = $false
	$mgText.SetBounds(0, 62, 1000, 44)
	$mgText.Anchor = 'Top,Left,Right'
	$mg.Controls.AddRange(@($mgRun, $mgDb, $mgInfoBtn, $mgText))
	$script:UI.MgPanel = $mg
	$script:UI.MgButtons = @($mgRun, $mgDb, $mgInfoBtn)

	# ---- operation tools (Scales Maintenance, Remote Lane Tools): toolbar + check list each
	$script:UI.Ops = @{ }
	foreach ($k in $script:OpsTools.Keys) { $script:UI.Ops[$k] = New_OpsToolControls $k; Fill_OpsList $k }

	# ---- connection card
	$conn = New_Panel 'Top' $T.Card 72
	$conn.Padding = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
	$dot = New_Label ([string][char]0x25CF) (New-Object System.Drawing.Font('Segoe UI', 22)) $T.Muted
	$dot.Location = New-Object System.Drawing.Point(14, 12)
	$connTitle = New_Label 'Detecting database...' $T.FontH $T.Text
	$connTitle.Location = New-Object System.Drawing.Point(54, 13)
	$connSub = New_Label '' $T.Font $T.Muted
	$connSub.Location = New-Object System.Drawing.Point(56, 42)
	$reconnect = New_Button 'Reconnect' 110
	$reconnect.Anchor = 'Top,Right'
	$reconnect.Add_Click({ Start_Detection -Reconnect })
	$conn.Controls.Add($dot)
	$conn.Controls.Add($connTitle)
	$conn.Controls.Add($connSub)
	$conn.Controls.Add($reconnect)
	$conn.Add_Resize({ $script:UI.Reconnect.Location = New-Object System.Drawing.Point(($this.Width - 126), 20) })
	$script:UI.Dot = $dot
	$script:UI.ConnTitle = $connTitle
	$script:UI.ConnSub = $connSub
	$script:UI.Reconnect = $reconnect

	$content.Controls.Add($toolbar)
	$content.Controls.Add($mg)
	foreach ($o in $script:UI.Ops.Values) { $content.Controls.Add($o.Toolbar) }
	$content.Controls.Add($conn)

	# ---- query list (Left of body)
	$listCard = New_Panel 'Left' $T.Card 0 360
	$listCard.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10)
	$list = New-Object System.Windows.Forms.CheckedListBox
	$list.Dock = 'Fill'
	$list.BackColor = $T.Card
	$list.ForeColor = $T.Text
	$list.BorderStyle = 'None'
	$list.CheckOnClick = $true
	$list.IntegralHeight = $false
	$list.Font = $T.Font
	$list.HorizontalScrollbar = $true
	$list.Add_DoubleClick({ Show_StepPreview })
	$listHead = New_Panel 'Top' $T.Card 34
	$listTitle = New_Label 'QUERIES' (New-Object System.Drawing.Font('Segoe UI Semibold', 8.5)) $T.Muted
	$listTitle.Location = New-Object System.Drawing.Point(0, 8)
	$all = New_Button 'All' 52 'ghost'
	$all.Height = 26
	$all.Anchor = 'Top,Right'
	$none = New_Button 'None' 58 'ghost'
	$none.Height = 26
	$none.Anchor = 'Top,Right'
	$script:UI.All = $all
	$script:UI.None = $none
	$all.Add_Click({ for ($i = 0; $i -lt $script:UI.List.Items.Count; $i++) { if (-not $script:Steps[$i].Danger) { $script:UI.List.SetItemChecked($i, $true) } } })
	$none.Add_Click({ for ($i = 0; $i -lt $script:UI.List.Items.Count; $i++) { $script:UI.List.SetItemChecked($i, $false) } })
	$listHead.Controls.Add($listTitle)
	$listHead.Controls.Add($all)
	$listHead.Controls.Add($none)
	$listHead.Add_Resize({ $script:UI.All.Location = New-Object System.Drawing.Point(($this.Width - 116), 2); $script:UI.None.Location = New-Object System.Drawing.Point(($this.Width - 58), 2) })
	$hint = New_Label 'Double-click a query to preview its SQL' (New-Object System.Drawing.Font('Segoe UI', 8.5)) $T.Muted
	$hint.AutoSize = $false
	$hint.Dock = 'Bottom'
	$hint.Height = 22
	$listCard.Controls.Add($list)
	$listCard.Controls.Add($hint)
	$listCard.Controls.Add($listHead)
	$script:UI.List = $list
	$script:UI.ListTitle = $listTitle
	$script:UI.ListCard = $listCard

	$gap = New_Panel 'Left' $T.Bg 0 14
	$script:UI.ListGap = $gap


	# ---- terminal (Fill of body)
	$termCard = New_Panel 'Fill' $T.TermBg
	$termCard.Padding = New-Object System.Windows.Forms.Padding(12, 0, 4, 8)
	$termHead = New_Panel 'Top' $T.TermBg 36
	$termTitle = New_Label 'TERMINAL' (New-Object System.Drawing.Font('Segoe UI Semibold', 8.5)) $T.Muted
	$termTitle.Location = New-Object System.Drawing.Point(0, 11)
	$termHead.Controls.Add($termTitle)
	$x = 0
	$script:UI.TermButtons = @()
	foreach ($spec in @(@('log', 'Open logs'), @('copy', 'Copy'), @('clear', 'Clear')))
	{
		$b = New_Button $spec[1] 84 'ghost'
		$b.Height = 26
		$b.Tag = $spec[0]
		$b.BackColor = $T.TermBg
		$b.Anchor = 'Top,Right'
		$b.Add_Click({
				switch ($this.Tag)
				{
					'clear' { $script:UI.Term.Clear() }
					'copy' { if ($script:UI.Term.TextLength) { [System.Windows.Forms.Clipboard]::SetText($script:UI.Term.Text) } }
					'log' { Start-Process -FilePath explorer.exe -ArgumentList ('"{0}"' -f $script:Paths.Log) }
				}
			})
		$termHead.Controls.Add($b)
		$script:UI.TermButtons += $b
	}
	$termHead.Add_Resize({ $x = $this.Width - 4; foreach ($b in $script:UI.TermButtons) { $x -= 88; $b.Location = New-Object System.Drawing.Point($x, 5) } })
	$term = New-Object System.Windows.Forms.RichTextBox
	$term.Dock = 'Fill'
	$term.ReadOnly = $true
	$term.BackColor = $T.TermBg
	$term.ForeColor = $T.TermText
	$term.BorderStyle = 'None'
	$term.Font = $T.FontMono
	$term.DetectUrls = $false
	$term.WordWrap = $true
	$term.HideSelection = $false
	# ---- prompt bar (Mini Ghost questions), hidden until Read_Prompt shows it
	$prompt = New_Panel 'Bottom' $T.Card 64
	$prompt.Padding = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
	$prompt.Visible = $false
	$pLabel = New_Label '' $T.FontBold $T.AccentHi
	$pLabel.Location = New-Object System.Drawing.Point(10, 8)
	$pBox = New-Object System.Windows.Forms.TextBox
	$pBox.BackColor = $T.TermBg
	$pBox.ForeColor = $T.Text
	$pBox.BorderStyle = 'FixedSingle'
	$pBox.Font = $T.FontMono
	$pBox.SetBounds(10, 30, 260, 24)
	$pBox.Add_KeyDown({
			if ($_.KeyCode -eq 'Enter') { $script:UI.PromptAnswer = $this.Text; $_.SuppressKeyPress = $true }
			elseif ($_.KeyCode -eq 'Escape') { $script:UI.PromptCancel = $true; $_.SuppressKeyPress = $true }
		})
	$pOk = New_Button 'OK' 70 'accent'
	$pOk.Height = 26
	$pOk.Location = New-Object System.Drawing.Point(280, 29)
	$pOk.Add_Click({ $script:UI.PromptAnswer = $script:UI.PromptBox.Text })
	$pCancel = New_Button 'Cancel' 76
	$pCancel.Height = 26
	$pCancel.Location = New-Object System.Drawing.Point(356, 29)
	$pCancel.Add_Click({ $script:UI.PromptCancel = $true })
	$pHint = New_Label '' (New-Object System.Drawing.Font('Segoe UI', 8.5)) $T.Muted
	$pHint.Location = New-Object System.Drawing.Point(444, 34)
	$prompt.Controls.AddRange(@($pLabel, $pBox, $pOk, $pCancel, $pHint))
	$script:UI.PromptBar = $prompt
	$script:UI.PromptLabel = $pLabel
	$script:UI.PromptBox = $pBox
	$script:UI.PromptHint = $pHint

	$termCard.Controls.Add($term)
	$termCard.Controls.Add($prompt)
	$termCard.Controls.Add($termHead)
	$script:UI.Term = $term

	$body.Controls.Add($termCard)
	$body.Controls.Add($gap)
	$body.Controls.Add($listCard)
	foreach ($o in $script:UI.Ops.Values) { $body.Controls.Add($o.Gap); $body.Controls.Add($o.Card) }

	# ---- timers: run output pump + detection poll
	$pump = New-Object System.Windows.Forms.Timer
	$pump.Interval = 100
	$pump.Add_Tick({ Update_RunOutput })
	$script:UI.Pump = $pump
	$mgPump = New-Object System.Windows.Forms.Timer
	$mgPump.Interval = 150
	$mgPump.Add_Tick({ Update_MgJobOutput })
	$script:UI.MgPump = $mgPump

	$f.Add_FormClosing({
			param ($s, $e)
			if ($script:MgJob)
			{
				if (-not (Show_Dialog 'Mini Ghost is running' "Closing now can leave this lane half-configured (name, INIs, SQL).`r`n`r`nClose anyway?" 'Close' 'Keep running' -Danger)) { $e.Cancel = $true; return }
			}
			$script:UI.Closing = $true
			if ($script:Run -and -not $script:Run.Sync.Done)
			{
				if (-not (Show_Dialog 'A query is still running' "Closing stops the running query (SQL rolls back the statement in progress).`r`n`r`nClose anyway?" 'Close' 'Keep running' -Danger)) { $e.Cancel = $true; return }
				Stop_Run
			}
			Save_Checked
			foreach ($k in $script:OpsTools.Keys) { Save_OpsChecked $k }
			Save_Settings
		})
	$f
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Switch_Maximize - borderless window: fill the work area (taskbar stays visible) / restore
# ---------------------------------------------------------------------------------------------------
function Switch_Maximize
{
	$f = $script:UI.Form
	if ($script:UI.NormalBounds)
	{
		$f.Bounds = $script:UI.NormalBounds
		$script:UI.NormalBounds = $null
	}
	else
	{
		$script:UI.NormalBounds = $f.Bounds
		$f.Bounds = [System.Windows.Forms.Screen]::FromControl($f).WorkingArea
	}
}

# ===================================================================================================
#                                    SECTION: Terminal + status
# ===================================================================================================

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Write_Term - one coloured line (or block) in the terminal + the file log
#   Kind: text | info | ok | warn | fail | step | sql | msg | table | muted
# ---------------------------------------------------------------------------------------------------
function Write_Term
{
	param ([string]$Text, [string]$Kind = 'text', [switch]$NoTime)
	$Text = Protect_Text $Text
	$tag = switch ($Kind) { 'ok' { '[ OK ] ' } 'warn' { '[WARN] ' } 'fail' { '[FAIL] ' } 'info' { '[INFO] ' } 'step' { '> ' } default { '' } }
	Write_LogFile ($tag + $Text)
	$rtb = $script:UI.Term
	if (-not $rtb -or $rtb.IsDisposed) { Write-Host ($tag + $Text); return }

	$color = switch ($Kind)
	{
		'ok' { $script:T.Ok } 'warn' { $script:T.Warn } 'fail' { $script:T.Fail } 'info' { $script:T.Info }
		'step' { $script:T.AccentHi } 'sql' { $script:T.Sql } 'muted' { $script:T.Muted } 'table' { $script:T.TermText }
		'red' { $script:T.Fail } 'yellow' { $script:T.Warn } 'blue' { $script:T.Info }   # coloured, no tag (reports)
		default { $script:T.TermText }
	}
	# keep long sessions responsive: drop the oldest quarter past ~1M characters (file log keeps all)
	if ($rtb.TextLength -gt 1000000)
	{
		$cut = $rtb.Text.IndexOf("`n", 250000)
		if ($cut -gt 0) { $rtb.ReadOnly = $false; $rtb.Select(0, $cut + 1); $rtb.SelectedText = ''; $rtb.ReadOnly = $true }
	}
	$prefix = if ($NoTime -or $Kind -in @('sql', 'table')) { '' } else { '{0:HH:mm:ss}  ' -f [datetime]::Now }
	if ($Kind -eq 'step' -and $rtb.TextLength) { $rtb.AppendText("`n") }
	$rtb.SelectionStart = $rtb.TextLength
	$rtb.SelectionLength = 0
	if ($prefix) { $rtb.SelectionColor = $script:T.Muted; $rtb.AppendText($prefix) }
	$rtb.SelectionColor = $color
	if ($Kind -eq 'sql') { $Text = (($Text -split "`r?`n") | ForEach-Object { '    ' + $_ }) -join "`n" }
	$rtb.AppendText($tag + $Text + "`n")
	$rtb.SelectionColor = $rtb.ForeColor
	$rtb.ScrollToCaret()
}

function Set_Status
{
	param ([string]$Text, [double]$Percent = -1)
	if ($script:UI.StatusText) { $script:UI.StatusText.Text = $Text }
	if ($Percent -ge 0 -and $script:UI.BarFill) { $script:UI.BarFill.Width = [int]($script:UI.BarInner.Width * [Math]::Min(100, $Percent) / 100) }
}

# ===================================================================================================
#                                  SECTION: Detection (startup)
# ===================================================================================================

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Start_Detection - runs $DetectionScript in a background runspace
# ---------------------------------------------------------------------------------------------------
function Start_Detection
{
	param ([switch]$Reconnect)
	if ($script:Run -and -not $script:Run.Sync.Done) { return }
	$sync = [hashtable]::Synchronized(@{ Stage = 'Starting...' })
	$ps = [PowerShell]::Create()
	$null = $ps.AddScript($script:DetectionScript).AddArgument($sync).AddArgument($script:Cfg)
	$script:Detect = @{ PS = $ps; Handle = $ps.BeginInvoke(); Sync = $sync; Started = [datetime]::Now }
	if ($Reconnect)
	{
		$script:UI.Dot.ForeColor = $script:T.Muted
		$script:UI.ConnTitle.Text = 'Reconnecting...'
		$script:UI.ConnSub.Text = ''
		$script:UI.Reconnect.Enabled = $false
		$script:UI.Run.Enabled = $false
		Write_Term 'Reconnecting...' 'info'
		$t = New-Object System.Windows.Forms.Timer
		$t.Interval = 100
		$t.Add_Tick({
				if ($script:Detect.Handle.IsCompleted) { $this.Stop(); $this.Dispose(); Complete_Detection }
				else { Set_Status $script:Detect.Sync.Stage }
			})
		$t.Start()
	}
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Complete_Detection - reads the result, fills the connection card, writes the terminal
# ---------------------------------------------------------------------------------------------------
function Complete_Detection
{
	$d = $null
	try { $d = @($script:Detect.PS.EndInvoke($script:Detect.Handle))[0] }
	catch { $d = @{ Ok = $false; Error = $_.Exception.Message } }
	finally { $script:Detect.PS.Dispose() }
	$script:Db = $d
	$secs = ([datetime]::Now - $script:Detect.Started).TotalSeconds

	$script:UI.Reconnect.Enabled = $true
	if ($d.Ok)
	{
		$verName = switch -regex ([string]$d.SqlVersion) { '^17\.' { 'SQL Server 2025' } '^16\.' { 'SQL Server 2022' } '^15\.' { 'SQL Server 2019' } '^14\.' { 'SQL Server 2017' } '^13\.' { 'SQL Server 2016' } '^12\.' { 'SQL Server 2014' } '^11\.' { 'SQL Server 2012' } '^10\.5' { 'SQL Server 2008 R2' } '^10\.' { 'SQL Server 2008' } default { 'SQL Server ' + $d.SqlVersion } }
		$edition = ([string]$d.Edition -replace '\s*\(64-bit\)', '')
		$server = if ($d.Server -eq '.') { $d.Machine } else { $d.Server -replace '^\.\\', "$($d.Machine)\" }
		$roleText = if ($d.Role -eq 'Lane') { "Lane $($d.Terminal)" } else { 'Store server' + $(if ($d.Terminal) { " ($($d.Terminal))" } else { '' }) }
		$script:UI.ConnTitle.Text = '{0}  on  {1}' -f $d.Database, $server
		$script:UI.ConnSub.Text = '{0}   |   Store {1}   |   {2} {3}   |   {4:N0} MB   |   {5}' -f $roleText, $(if ($d.Store) { $d.Store } else { '?' }), $verName, $edition, $d.SizeMB, $d.State
		$healthy = ($d.State -eq 'ONLINE')
		$script:UI.Dot.ForeColor = if ($healthy) { $script:T.Ok } else { $script:T.Warn }

		Write_Term ('Machine {0} detected as {1}.' -f $d.Machine, $roleText) 'info'
		if ($d.Storeman) { Write_Term ('Storeman folder: {0} (via {1})' -f $d.Storeman, $d.Source) 'muted' }
		foreach ($n in $d.Notes) { Write_Term $n 'muted' }
		Write_Term ('Database {0} connection successful  ({1}, {2} {3}, {4:0.0} s)' -f $d.Database, $server, $verName, $edition, $secs) 'ok'
		if (-not $healthy) { Write_Term ("Database {0} is {1} - use Repair mode." -f $d.Database, $d.State) 'warn' }
		if ($d.IsSysadmin -eq $false) { Write_Term ('{0} is not a SQL sysadmin - some queries (memory, logins, repair) will fail. Run the toolkit as Administrator.' -f [Environment]::UserName) 'warn' }
		$script:UI.Run.Enabled = $true
		Select_Mode $(if ($healthy) { $script:Mode } else { 'Repair' }) -Force
		Set_Status ('Connected to {0}' -f $d.Database) 0
	}
	else
	{
		$script:UI.Dot.ForeColor = $script:T.Fail
		$script:UI.ConnTitle.Text = 'No database connection'
		$script:UI.ConnSub.Text = [string]$d.Error
		foreach ($n in $d.Notes) { Write_Term $n 'muted' }
		Write_Term ('Database connection failed: {0}' -f $d.Error) 'fail'
		Write_Term 'Check that SQL Server is running on this machine, then press Reconnect.' 'muted'
		$script:UI.Run.Enabled = $false
		Select_Mode $script:Mode -Force
		Set_Status 'Not connected'
	}
	$script:MgInfo = $null
	Update_MiniGhostNav
}

# ===================================================================================================
#                                SECTION: Database Maintenance tool
# ===================================================================================================

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Select_Mode - Maintenance / Repair / Custom (.sqi): rebuilds the query list
# ---------------------------------------------------------------------------------------------------
function Select_Mode
{
	param ([string]$Mode, [switch]$Force)
	if ($script:Run -and -not $script:Run.Sync.Done) { return }
	$file = $null
	if ($Mode -eq 'Custom')
	{
		$dlg = New-Object System.Windows.Forms.OpenFileDialog
		$dlg.Filter = 'SQL scripts (*.sqi;*.sql)|*.sqi;*.sql|All files (*.*)|*.*'
		$dlg.Title = 'Load a SQL script'
		if ($script:CustomFile) { $dlg.InitialDirectory = Split-Path $script:CustomFile }
		if ($Force -and $script:CustomFile) { $file = $script:CustomFile }
		elseif ($dlg.ShowDialog($script:UI.Form) -eq 'OK') { $file = $dlg.FileName }
		else { return }
		$script:CustomFile = $file
	}
	if (-not $Force) { Save_Checked }
	$script:Mode = $Mode

	foreach ($k in $script:UI.ModeButtons.Keys)
	{
		$b = $script:UI.ModeButtons[$k]
		if ($k -eq $Mode) { $b.BackColor = $script:T.Accent; $b.ForeColor = [System.Drawing.Color]::White; $b.FlatAppearance.MouseOverBackColor = $script:T.AccentHi }
		else { $b.BackColor = $script:T.Card; $b.ForeColor = $script:T.Text; $b.FlatAppearance.MouseOverBackColor = $script:T.CardHi }
	}

	try { $script:Steps = @(Get_ModeSteps $Mode $file) }
	catch { $script:Steps = @(); Write_Term ('Could not load the queries: {0}' -f $_.Exception.Message) 'fail' }

	$saved = $script:Settings[(Get_SettingsKey)]
	$list = $script:UI.List
	$list.BeginUpdate()
	$list.Items.Clear()
	foreach ($s in $script:Steps)
	{
		$label = $s.Name
		if ($s.Danger) { $label = '! ' + $label }
		$checked = if ($saved -and $saved.ContainsKey($s.Name) -and -not $s.Danger) { $saved[$s.Name] } else { [bool]$s.Default }
		[void]$list.Items.Add($label, $checked)
	}
	$list.EndUpdate()

	$what = switch ($Mode) { 'Maintenance' { 'maintenance' } 'Repair' { 'repair' } default { Split-Path $file -Leaf } }
	$script:UI.ListTitle.Text = ('QUERIES  -  {0}  ({1})' -f $what.ToUpper(), $script:Steps.Count)
	if (-not $Force -or $Mode -eq 'Custom')
	{
		$roleWord = if ($script:Db -and $script:Db.Role) { $script:Db.Role.ToLower() } else { 'store' }
		$src = switch ($Mode) { 'Maintenance' { "$roleWord maintenance script" } 'Repair' { 'guarded repair steps' } default { $file } }
		Write_Term ('{0} mode: {1} queries loaded ({2}).' -f $Mode, $script:Steps.Count, $src) 'info'
	}
}

function Get_SettingsKey
{
	$role = if ($script:Db -and $script:Db.Role) { $script:Db.Role } else { 'Unknown' }
	$m = if ($script:Mode -eq 'Custom' -and $script:CustomFile) { 'Custom:' + (Split-Path $script:CustomFile -Leaf) } else { $script:Mode }
	"$role|$m"
}

# Remember which queries are ticked (per role + mode); dangerous steps are never remembered
function Save_Checked
{
	if (-not $script:UI.List -or $script:Steps.Count -eq 0) { return }
	$h = @{ }
	for ($i = 0; $i -lt $script:Steps.Count; $i++)
	{
		if ($script:Steps[$i].Danger) { continue }
		$h[$script:Steps[$i].Name] = $script:UI.List.GetItemChecked($i)
	}
	$script:Settings[(Get_SettingsKey)] = $h
}

# Double-click: show the selected query's SQL in the terminal
function Show_StepPreview
{
	$i = $script:UI.List.SelectedIndex
	if ($i -lt 0) { return }
	$s = $script:Steps[$i]
	Write_Term ('Preview: {0}{1}' -f $s.Name, $(if ($s.Master) { '  (runs on master)' } else { '' })) 'step'
	Write_Term $s.Sql 'sql'
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Start_SelectedRun - confirmations, then the run engine in a background runspace
# ---------------------------------------------------------------------------------------------------
function Start_SelectedRun
{
	if (-not ($script:Db -and $script:Db.Ok)) { Write_Term 'Not connected to a database.' 'fail'; return }
	if ($script:Run -and -not $script:Run.Sync.Done) { return }

	$sel = @()
	for ($i = 0; $i -lt $script:Steps.Count; $i++) { if ($script:UI.List.GetItemChecked($i)) { $sel += $script:Steps[$i] } }
	if ($sel.Count -eq 0) { Write_Term 'No queries selected.' 'warn'; return }
	Save_Checked
	Save_Settings

	$db = $script:Db.Database
	$names = ($sel | ForEach-Object { '  - ' + $_.Name }) -join "`r`n"
	if ($sel.Count -gt 8) { $names = (($sel | Select-Object -First 8 | ForEach-Object { '  - ' + $_.Name }) -join "`r`n") + "`r`n  ... and $($sel.Count - 8) more" }

	$danger = @($sel | Where-Object { $_.Danger })
	$single = @($sel | Where-Object { $_.SingleUser })
	if ($danger.Count -gt 0)
	{
		$noBackup = -not ($sel | Where-Object { $_.Backup })
		$msg = "This can DELETE DATA from $db to make it usable again:`r`n" + (($danger | ForEach-Object { '  - ' + $_.Name }) -join "`r`n") +
		"`r`n`r`nClose SMS on this machine first - everyone is disconnected from $db." +
		$(if ($noBackup) { "`r`nNO BACKUP IS SELECTED." } else { '' }) + "`r`n`r`nType REPAIR to continue:"
		if (-not (Show_Dialog 'Data-loss repair' $msg 'Run repair' 'Cancel' -TypeToConfirm 'REPAIR' -Danger)) { Write_Term 'Cancelled.' 'warn'; return }
	}
	elseif ($single.Count -gt 0)
	{
		if (-not (Show_Dialog 'Repair' "Run $($sel.Count) step(s) on $db?`r`n$names`r`n`r`nThe repair disconnects everyone from $db (single-user mode). Close SMS on this machine first." 'Run' 'Cancel')) { Write_Term 'Cancelled.' 'warn'; return }
	}
	else
	{
		if (-not (Show_Dialog 'Run queries' "Run $($sel.Count) quer$(if ($sel.Count -eq 1) { 'y' } else { 'ies' }) on $db ($($script:Db.Role.ToLower()))?`r`n$names" 'Run' 'Cancel')) { Write_Term 'Cancelled.' 'warn'; return }
	}

	$restore = Expand_Sql ($script:RepairSteps | Where-Object { $_.Name -like 'Return the database*' } | Select-Object -First 1).Sql
	$sync = [hashtable]::Synchronized(@{
			Queue = New-Object System.Collections.Concurrent.ConcurrentQueue[object]
			Cancel = $false; Done = $false; Index = 0; Current = ''; StepStart = [datetime]::Now; Cmd = $null
			ShowSql = $script:UI.ShowSql.Checked
		})
	$steps = @($sel | ForEach-Object { @{ Name = $_.Name; Sql = $_.Sql; Master = [bool]$_.Master; SingleUser = [bool]$_.SingleUser; Danger = [bool]$_.Danger; Backup = [bool]$_.Backup } })
	$ps = [PowerShell]::Create()
	$null = $ps.AddScript($script:RunEngineScript).AddArgument($sync).AddArgument($script:Db.ConnectionString).AddArgument($script:Db.MasterConnectionString).
	AddArgument($steps).AddArgument($script:Cfg.SqlPass).AddArgument($restore)

	Write_Term ('==== {0} on {1}: {2} quer{3} ====' -f $script:Mode, $db, $steps.Count, $(if ($steps.Count -eq 1) { 'y' } else { 'ies' })) 'info'
	$script:Run = @{ PS = $ps; Handle = $ps.BeginInvoke(); Sync = $sync; Started = [datetime]::Now; Count = $steps.Count }
	Set_RunningUi $true
	$script:UI.Pump.Start()
}

function Stop_Run
{
	if (-not $script:Run -or $script:Run.Sync.Done) { return }
	$script:Run.Sync.Cancel = $true
	try { if ($script:Run.Sync.Cmd) { $script:Run.Sync.Cmd.Cancel() } }
	catch { }
	Write_Term 'Stopping after the current statement...' 'warn'
	$script:UI.Stop.Enabled = $false
}

function Set_RunningUi
{
	param ([bool]$Running)
	$script:UI.Run.Enabled = (-not $Running) -and [bool]($script:Db -and $script:Db.Ok)
	$script:UI.Stop.Enabled = $Running -and [bool]$script:Run -and -not $script:Run.Sync.Done
	$script:UI.Reconnect.Enabled = -not $Running
	$script:UI.List.Enabled = -not $Running
	foreach ($b in $script:UI.ModeButtons.Values) { $b.Enabled = -not $Running }
	foreach ($b in $script:UI.MgButtons) { $b.Enabled = -not $Running }
	foreach ($o in $script:UI.Ops.Values) { $o.Run.Enabled = -not $Running; $o.Info.Enabled = -not $Running; $o.List.Enabled = -not $Running }
	if ($script:UI.SmsFn) { $script:UI.SmsFn.Enabled = -not $Running }
	$script:UI.NavDb.Enabled = -not $Running
	if ($Running) { foreach ($k in $script:UI.Nav.Keys) { if ($k -ne 'Db') { $script:UI.Nav[$k].Button.Enabled = $false } } } else { Update_MiniGhostNav }
}

# ---------------------------------------------------------------------------------------------------
# FUNCTION: Update_RunOutput - timer tick: drains the run queue into the terminal, moves the bar
# ---------------------------------------------------------------------------------------------------
function Update_RunOutput
{
	$r = $script:Run
	if (-not $r) { $script:UI.Pump.Stop(); return }
	$item = $null
	$n = 0
	while ($n -lt 200 -and $r.Sync.Queue.TryDequeue([ref]$item))
	{
		$n++
		if ($item.T -eq 'done')
		{
			$parts = $item.X.Split('|', 2)
			Write_Term $parts[1] $parts[0]
		}
		elseif ($item.T -eq 'skip') { Write_Term $item.X 'muted' }
		else { Write_Term $item.X $item.T }
	}
	$cur = [datetime]::Now - $r.Sync.StepStart
	$total = [datetime]::Now - $r.Started
	$script:UI.Elapsed.Text = '{0:hh\:mm\:ss}' -f $total
	$pct = if ($r.Count) { 100.0 * $r.Sync.Index / $r.Count } else { 0 }
	if (-not $r.Sync.Done) { Set_Status ('Running {0}/{1}: {2}  ({3:mm\:ss})' -f ($r.Sync.Index + 1), $r.Count, $r.Sync.Current, $cur) $pct }

	if ($r.Sync.Done -and $r.Sync.Queue.IsEmpty)
	{
		$script:UI.Pump.Stop()
		try { $r.PS.EndInvoke($r.Handle) | Out-Null }
		catch { Write_Term ('Run engine error: {0}' -f $_.Exception.Message) 'fail' }
		$r.PS.Dispose()
		Set_Status 'Finished' 100
		Set_RunningUi $false
	}
}

# ===================================================================================================
#                                        SECTION: Startup
# ---------------------------------------------------------------------------------------------------
#   1. WinForms + splash (shown at once)  2. detection starts in the background  3. main window is
#   built meanwhile  4. wait for detection with the bar moving  5. close splash, show main window.
# ===================================================================================================

function Hide_ConsoleWindow
{
	# Only when started by double-click / "Run with PowerShell" (own console), never a typed-in terminal
	try
	{
		$parent = (Get-CimInstance Win32_Process -Filter "ProcessId=$PID" -ErrorAction Stop).ParentProcessId
		$pname = (Get-Process -Id $parent -ErrorAction Stop).ProcessName
		if ($pname -notin @('explorer', 'cmd')) { return }
		Add-Type -Namespace SmsToolkit -Name Win32 -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@
		$h = [SmsToolkit.Win32]::GetConsoleWindow()
		if ($h -ne [System.IntPtr]::Zero) { $null = [SmsToolkit.Win32]::ShowWindow($h, 0) }
	}
	catch { }
}

function Start_App
{
	Add-Type -AssemblyName System.Windows.Forms, System.Drawing
	[System.Windows.Forms.Application]::EnableVisualStyles()
	Initialize_Theme
	Initialize_Paths

	$splash = $null
	if (-not $SelfTest)
	{
		$splash = New_Splash
		$splash.Form.Show()
		Update_Splash $splash 8 'Starting...'
	}

	# Background: database detection (+ hiding the console, which compiles a small type)
	Start_Detection
	$hider = $null
	if (-not $SelfTest -and -not $KeepConsole)
	{
		$hider = [PowerShell]::Create()
		$null = $hider.AddScript(${function:Hide_ConsoleWindow}.ToString())
		$hiderHandle = $hider.BeginInvoke()
	}

	# Foreground meanwhile: settings + main window
	Update_Splash $splash 25 'Loading settings...'
	Load_Settings
	Update_Splash $splash 45 'Building the interface...'
	$form = New_MainWindow
	Update_Splash $splash 65 $script:Detect.Sync.Stage

	# Wait for detection (max 30 s), the bar creeping toward 95 %
	$p = 65.0
	while (-not $script:Detect.Handle.IsCompleted -and ([datetime]::Now - $script:Detect.Started).TotalSeconds -lt 30)
	{
		$p = [Math]::Min(95, $p + 0.4)
		if ($splash) { Update_Splash $splash $p $script:Detect.Sync.Stage }
		else { Start-Sleep -Milliseconds 20 }
	}

	Write_Term ('{0} v{1}  -  {2}' -f $script:App.Name, $script:App.Version, $script:App.Credits) 'muted' -NoTime
	if ($script:Detect.Handle.IsCompleted) { Complete_Detection }
	else
	{
		$script:Db = @{ Ok = $false; Error = 'detection timed out after 30 s' }
		Write_Term 'Database detection timed out after 30 s - press Reconnect.' 'fail'
		Select_Mode $script:Mode -Force
	}

	if ($SelfTest)
	{
		Write-Host ''
		Write-Host '---- SELF-TEST ----'
		$script:Db.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ('{0,-24} {1}' -f $_.Name, (Protect_Text ([string]($_.Value -join '; ')))) }
		foreach ($m in @('Maintenance', 'Repair'))
		{
			$script:Mode = $m
			$st = @(Get_ModeSteps $m)
			Write-Host ('{0,-24} {1} steps: {2}' -f "$m steps", $st.Count, (($st | ForEach-Object { $_.Name }) -join ' | '))
		}
		$lane = $script:Db; $script:Db = @{ Role = 'Lane'; Database = 'LANESQL'; Store = '123' }
		$ls = @(Get_ModeSteps 'Maintenance')
		Write-Host ('{0,-24} {1} steps; drop-tables SQL: {2}' -f 'Lane maintenance', $ls.Count, (($ls | Where-Object { $_.Name -like 'Drop*' }).Sql -split "`n" | Where-Object { $_ -match 'LIKE' }).Trim())
		$script:Db = @{ Role = 'Server'; Database = 'STORESQL'; Store = '' }
		$ss = @(Get_ModeSteps 'Maintenance')
		Write-Host ('{0,-24} {1}' -f 'No-store drop-tables', (($ss | Where-Object { $_.Name -like 'Drop*' }).Sql -split "`n" | Where-Object { $_ -match 'LIKE' }).Trim())
		Write-Host ('{0,-24} {1}' -f 'Password masked', ((Protect_Text (Expand_Sql '{{SQLPASS}}')) -eq '********'))
		Write-Host ('{0,-24} {1}' -f 'Log folder', $script:Paths.Log)

		# ---- Mini Ghost core against a temporary fake storeman (no SQL / rename / IP / registry)
		$fx = Join-Path $env:TEMP ('SMS-Toolkit-fixture-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
		$null = New-Item -Path (Join-Path $fx 'Office\Dbs') -ItemType Directory -Force
		[System.IO.File]::WriteAllText((Join-Path $fx 'Startup.ini'), "[SMS]`r`nSTORE=0231`r`nTER=005`r`nDBSERVER=POS005\SQLEXPRESS`r`nDBNAME=LANESQL`r`nREDIRMAIL=0231901`r`nREDIRMSG=0231901`r`n", [System.Text.Encoding]::GetEncoding(1252))
		$null = New-Item -Path (Join-Path $fx 'Office\XF0231005') -ItemType Directory -Force
		$null = New-Item -Path (Join-Path $fx 'Office\XF0231007') -ItemType Directory -Force
		$null = New-Item -Path (Join-Path $fx 'Office\XF0231003') -ItemType Directory -Force
		Initialize_MiniGhost_Context -StoremanRoot $fx
		Write-Host ('{0,-24} role {1}, store {2}' -f 'MG fixture', $script:NodeRole, (Get_Store_Number_From_INI))
		foreach ($cur in @('POS005', 'LN0231005', '0231LANE005', 'WEIRD-PC'))
		{
			$c = Get_Machine_Name_Context -MachineName '007' -CurrentMachineName $cur
			Write-Host ('{0,-24} {1} + lane 007 -> {2}' -f 'MG name', $cur, $(if ($c.Success) { "$($c.MachineName) (ter $($c.MachineNumber))" } else { 'ERROR: ' + ($c.ErrorMessage -split "`n")[0] }))
		}
		$ctx = Get_Machine_Name_Context -MachineName '007' -CurrentMachineName 'POS005'
		$ps = [PowerShell]::Create()
		$null = $ps.AddScript((Get_MgRunspaceScript)).AddArgument($fx).AddArgument('Update_Startup_Ini_For_Machine').AddArgument(@{ MachineContext = $ctx; StartupIniPath = (Join-Path $fx 'Startup.ini') })
		$r = @($ps.Invoke())
		Write-Host ('{0,-24} errors {1}, result: {2}' -f 'MG runspace', $ps.Streams.Error.Count, ($r | Select-Object -Last 1).Details)
		$ps.Dispose()
		$ps = [PowerShell]::Create()
		$null = $ps.AddScript((Get_MgRunspaceScript)).AddArgument($fx).AddArgument('Update_INIs').AddArgument(@{ newStoreNumber = '0231'; MachineName = 'POS007'; OldStoreNumber = '0231'; OldMachineName = 'POS005' })
		$r = @($ps.Invoke())
		Write-Host ('{0,-24} returned {1}, {2} log line(s), errors {3}' -f 'MG Update_INIs', ($r | Select-Object -Last 1), $ps.Streams.Information.Count, $ps.Streams.Error.Count)
		foreach ($er in $ps.Streams.Error) { Write-Host ('   ' + $er.Exception.Message) }
		$ps.Dispose()
		$ps = [PowerShell]::Create()
		$null = $ps.AddScript((Get_MgRunspaceScript)).AddArgument($fx).AddArgument('Remove_Old_XZ_Folders').AddArgument(@{ StoreNumber = '0231'; MachineName = 'POS007' })
		$null = $ps.Invoke()
		Write-Host ('{0,-24} left: {1}' -f 'MG X/Z folders', ((Get-ChildItem (Join-Path $fx 'Office') -Directory | ForEach-Object Name) -join ', '))
		$ps.Dispose()
		Write-Host ('{0,-24} {1}' -f 'MG Startup.ini after', ((Get-Content (Join-Path $fx 'Startup.ini')) -join ' | '))
		Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue

		# ---- Scales: Fix_Deploy_CHG against a temporary Office folder + fake ScaleManagementApp
		$sx = Join-Path $env:TEMP ('SMS-Toolkit-scales-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
		$null = New-Item -Path (Join-Path $sx 'Office') -ItemType Directory -Force
		$null = New-Item -Path (Join-Path $sx 'ScaleCommApp') -ItemType Directory -Force
		$fakeApp = Join-Path $sx 'ScaleCommApp\ScaleManagementApp.exe'
		[System.IO.File]::WriteAllText($fakeApp, 'x')
		$deploy = Join-Path $sx 'Office\DEPLOY_CHG.sql'
		[System.IO.File]::WriteAllText($deploy, "@WIZRPL(X=1);`r`n/* Deploy price changes to the scales */`r`n@EXEC(RUN='C:\ScaleCommApp\ScaleManagementApp.exe');`r`n`r`n")
		$script:Db = @{ Ok = $false; Storeman = $sx; Role = 'Server' }
		Initialize_ScalesContext
		$script:ScaleManagementAppFastDeployPathCandidates = @()
		$script:ScaleManagementAppFastDeployPath = Join-Path $sx 'none.exe'
		$script:ScaleManagementAppPathCandidates = @($fakeApp)
		$before = $script:UI.Term.TextLength
		$ok = Fix_Deploy_CHG -OfficePath (Join-Path $sx 'Office')
		Write-Host ('{0,-24} returned {1}; file: {2}' -f 'Scales Fix_Deploy_CHG', ($ok | Select-Object -Last 1), ((Get-Content $deploy) -join ' | '))
		Write-Host ('{0,-24} {1}' -f 'Scales terminal lines', (($script:UI.Term.Text.Substring($before).Trim() -split "`n" | ForEach-Object { $_.Trim() }) -join ' || '))
		Write-Host ('{0,-24} {1}' -f 'Scales deploy regex', $script:ScaleCommDeployExecRegex)
		Remove-Item -LiteralPath $sx -Recurse -Force -ErrorAction SilentlyContinue

		# ---- tool switching with a server role
		$script:Db = @{ Ok = $false; Role = 'Server'; Storeman = 'C:\Storeman' }
		Update_MiniGhostNav
		$navState = { ($script:UI.Nav.Keys | ForEach-Object { '{0}={1}' -f $_, $script:UI.Nav[$_].Button.Enabled }) -join ', ' }
		Write-Host ('{0,-24} {1}' -f 'Nav (server)', (& $navState))
		$script:Db = @{ Ok = $false; Role = 'Lane'; Storeman = 'C:\Storeman' }
		Update_MiniGhostNav
		Write-Host ('{0,-24} {1}' -f 'Nav (lane)', (& $navState))
		foreach ($k in $script:UI.Ops.Keys) { Write-Host ('{0,-24} {1}' -f "$k list", (($script:UI.Ops[$k].List.Items | ForEach-Object { $_ }) -join ' | ')) }

		# ---- Remote lane plumbing: mailslot type, SMS file writer
		Ensure_MailslotSender_Type
		$sent = [MailslotSender]::SendMailslotCommand('\\.\mailslot\SMS_Toolkit_SelfTest_None', '@exec(RESTART_ALL=PROGRAMS).')
		Write-Host ('{0,-24} type loaded, send to a missing slot -> {1} (expected False)' -f 'Mailslot sender', $sent)
		$sqi = Join-Path $env:TEMP 'SMS-Toolkit-selftest.sqi'
		Write_Sms_File -Path $sqi -Content "@dbEXEC(UPDATE SAL_HDR SET F1067 = 'CLOSE' WHERE F1067 <> 'CLOSE')`n"
		$b = [System.IO.File]::ReadAllBytes($sqi)
		Write-Host ('{0,-24} {1} bytes, ends CRLF={2}, archive bit={3}' -f 'Write_Sms_File', $b.Length, ($b[-2] -eq 13 -and $b[-1] -eq 10), ([bool]((Get-Item $sqi).Attributes -band [System.IO.FileAttributes]::Archive)))
		Remove-Item -LiteralPath $sqi -Force -ErrorAction SilentlyContinue

		# ---- Repair Windows: process streaming + read-only reports on this PC
		$before = $script:UI.Term.TextLength
		$r = Invoke_LiveProcess (Get_SystemToolPath 'cmd.exe') '/c echo first line& echo Verification 45% complete.& echo last line& exit 3'
		Write-Host ('{0,-24} exit {1}, lines kept: {2}' -f 'Live process', $r.Exit, ($r.Text -replace "`n", ' | '))
		$null = Write_DiskHealth
		$null = Write_EventReport -Days 2
		$out = $script:UI.Term.Text.Substring($before) -split "`n" | Where-Object { $_.Trim() }
		Write-Host ('{0,-24} {1} terminal lines; sample:' -f 'Disk + event report', $out.Count)
		$out | Select-Object -First 14 | ForEach-Object { Write-Host ('    ' + $_.TrimEnd()) }

		# ---- optional: render the splash + main window to PNG (set SMSTK_SNAPSHOT to a folder)
		if ($env:SMSTK_SNAPSHOT)
		{
			$snap = {
				param ($Form, $Path)
				$Form.StartPosition = 'Manual'
				$Form.Location = New-Object System.Drawing.Point(-6000, -6000)
				$Form.Show()
				[System.Windows.Forms.Application]::DoEvents()
				$bmp = New-Object System.Drawing.Bitmap($Form.Width, $Form.Height)
				$Form.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0, 0, $Form.Width, $Form.Height)))
				$bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
				$bmp.Dispose()
				$Form.Hide()
			}
			$sp = New_Splash
			$sp.Fill.Width = 300; $sp.Status.Text = 'Connecting to LANESQL on .\SQLEXPRESS...'
			& $snap $sp.Form (Join-Path $env:SMSTK_SNAPSHOT 'splash.png')
			$sp.Form.Dispose()
			$script:Db = @{ Ok = $false; Role = 'Server' }
			Update_MiniGhostNav
			& $snap $form (Join-Path $env:SMSTK_SNAPSHOT 'main.png')
			Write-Host ('{0,-24} {1}' -f 'Snapshots', $env:SMSTK_SNAPSHOT)
		}

		# ---- Read_Prompt must return exactly the typed string (a timer "types" the answer)
		$typer = New-Object System.Windows.Forms.Timer
		$typer.Interval = 150
		$typer.Add_Tick({ $script:UI.PromptAnswer = $script:SelfTestAnswer })
		$script:SelfTestAnswer = '3'
		$typer.Start()
		$ans = Read_Prompt 'Self-test question' '7' '^\d{1,3}$' 'digits'
		Write-Host ('{0,-24} type={1}, count={2}, value={3}' -f 'Read_Prompt', $(if ($null -ne $ans) { $ans.GetType().Name } else { 'null' }), @($ans).Count, ($ans -join ','))
		# the real Event Viewer step (prompt answered by the timer, then the report)
		$script:SelfTestAnswer = '1'
		$okEv = Invoke_RepairStep @{ Function = 'Events' }
		$typer.Stop(); $typer.Dispose()
		Write-Host ('{0,-24} returned {1}' -f 'Events step (real path)', $okEv)

		# ---- Scheduler scripts
		$script:Db = @{ Ok = $false; Role = 'Lane'; Database = 'LANESQL'; Store = '231' }
		$ls = Expand_Sql $script:SqlScheduleLane
		$script:Db = @{ Ok = $false; Role = 'Server'; Database = 'STORESQL'; Store = '231' }
		$ss = Expand_Sql $script:SqlScheduleServer
		Write-Host ('{0,-24} lane {1} chars (LANESQL={2}, store tables={3}), server {4} chars (STORESQL={5}), no placeholders left={6}' -f 'Schedule scripts', $ls.Length, ($ls -match 'SHRINKFILE \(LANESQL\)'), ($ls -match "M231%"), $ss.Length, ($ss -match 'SHRINKFILE \(STORESQL\)'), (-not ($ls + $ss -match '\{\{')))

		# ---- every tool by role
		foreach ($rl in @('Server', 'Lane', $null))
		{
			$script:Db = @{ Ok = $false; Role = $rl }
			Update_MiniGhostNav
			Write-Host ('{0,-24} {1}  | SMS functions button={2}' -f "Nav ($(if ($rl) { $rl } else { 'unknown' }))", (& $navState), $script:UI.SmsFn.Visible)
		}
		$form.Dispose()
		return
	}

	Update_Splash $splash 100 'Ready'
	$form.Add_Shown({ $this.Activate() })
	$splash.Form.Close()
	$splash.Form.Dispose()
	[System.Windows.Forms.Application]::Run($form)

	if ($hider) { try { $hider.Dispose() } catch { } }
}

Start_App
