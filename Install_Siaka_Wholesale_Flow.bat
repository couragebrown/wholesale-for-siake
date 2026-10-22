@echo off
title Siaka Wholesale Flow - Installer
cd /d "%~dp0"
cls

echo ==========================================================
echo    SIAKA WHOLESALE FLOW - PC INSTALLER & SETUP
echo ==========================================================
echo.
echo Installing Siaka Wholesale Flow to:
echo   %~dp0
echo.
echo Setting up local database, folders, auto-backups, and shortcuts...
echo.

set "CURRENT_DIR=%~dp0"
if "%CURRENT_DIR:~-1%"=="\" set "CURRENT_DIR=%CURRENT_DIR:~0,-1%"

powershell -ExecutionPolicy Bypass -NoProfile -File "%CURRENT_DIR%\installer.ps1" -TargetDir "%CURRENT_DIR%"

echo.
echo ==========================================================
echo   SETUP COMPLETE!
echo ==========================================================
echo.
set /p launchApp="Would you like to launch Siaka Wholesale Flow now? (Y/N): "
if /i "%launchApp%"=="Y" (
    start "" "%~dp0Launch_Wholesale_App.vbs"
)
exit
