@echo off
title Siaka Wholesale Flow - Starting...
cd /d "%~dp0"

echo ========================================================
echo   Siaka Wholesale Flow - Local Desktop App
echo   Saving data directly to your PC
echo ========================================================
echo.

:: Test if server is already responding
powershell -Command "try { $r = Invoke-WebRequest -Uri 'http://127.0.0.1:49221/api/data' -UseBasicParsing -TimeoutSec 1; exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Starting local background service...
    start /min powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0app\server.ps1" -Port 49221
    timeout /t 2 /nobreak >nul
) else (
    echo Background service is already active.
)

echo Opening Wholesale Flow window...
set "PROFILE_DIR=%~dp0app\data\app_profile"
if not exist "%PROFILE_DIR%" mkdir "%PROFILE_DIR%"

if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --user-data-dir="%PROFILE_DIR%" --app="http://127.0.0.1:49221"
) else (
    start "" "msedge.exe" --user-data-dir="%PROFILE_DIR%" --app="http://127.0.0.1:49221"
)

exit
