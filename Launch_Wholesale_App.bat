@echo off
title Siaka Wholesale Flow
cd /d "%~dp0"

:: 1. Ensure required local data folders exist on this drive
if not exist "%~dp0app\data" mkdir "%~dp0app\data"
if not exist "%~dp0app\data\backups" mkdir "%~dp0app\data\backups"
set "PROFILE_DIR=%~dp0app\data\app_profile"
if not exist "%PROFILE_DIR%" mkdir "%PROFILE_DIR%"

:: 2. Initialize clean database if not already created
if not exist "%~dp0app\data\wholesale_data.json" (
    if exist "%~dp0app\data\clean_schema_template.json" (
        copy /y "%~dp0app\data\clean_schema_template.json" "%~dp0app\data\wholesale_data.json" >nul
    ) else (
        powershell -ExecutionPolicy Bypass -Command "[System.IO.File]::WriteAllText('%~dp0app\data\wholesale_data.json', '{\"products\":[],\"dealers\":[],\"orders\":[],\"payments\":[],\"suppliers\":[],\"supplierPurchases\":[],\"supplierPayments\":[]}', [System.Text.Encoding]::UTF8)"
    )
)

:: 3. Test if local background server is already active on port 49221
powershell -ExecutionPolicy Bypass -Command "try { $r = Invoke-WebRequest -Uri 'http://127.0.0.1:49221/api/data' -UseBasicParsing -TimeoutSec 1; exit 0 } catch { exit 1 }" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    start /min powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0app\server.ps1" -Port 49221
    timeout /t 2 /nobreak >nul
)

:: 4. Launch Desktop Window (Edge app mode or fallback)
if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" (
    start "" "C:\Program Files\Microsoft\Edge\Application\msedge.exe" --user-data-dir="%PROFILE_DIR%" --app="http://127.0.0.1:49221"
) else if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --user-data-dir="%PROFILE_DIR%" --app="http://127.0.0.1:49221"
) else if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --user-data-dir="%PROFILE_DIR%" --app="http://127.0.0.1:49221"
) else if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --user-data-dir="%PROFILE_DIR%" --app="http://127.0.0.1:49221"
) else (
    start http://127.0.0.1:49221
)

exit
