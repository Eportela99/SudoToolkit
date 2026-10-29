@echo off
rem TechRoute launcher - always runs the latest version from GitHub (Eportela99/SudoToolkit)
title TechRoute
set "URL=https://raw.githubusercontent.com/Eportela99/SudoToolkit/main/Tecnica_System/TechRoute/TechRoute.ps1"
set "DEST=%TEMP%\SudoToolkit\TechRoute.ps1"

echo Downloading the latest TechRoute from GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol='Tls12'; $null = New-Item -ItemType Directory -Force (Split-Path $env:DEST); $tmp = $env:DEST + '.download'; Invoke-WebRequest -UseBasicParsing -Uri ($env:URL + '?t=' + [DateTime]::UtcNow.Ticks) -OutFile $tmp; Move-Item -Force $tmp $env:DEST"
if errorlevel 1 (
    if exist "%DEST%" (
        echo Download failed - starting the last downloaded copy.
        timeout /t 3 >nul
    ) else (
        echo.
        echo Download failed. Check the internet connection and try again.
        pause
        exit /b 1
    )
)

start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%DEST%"
exit /b 0