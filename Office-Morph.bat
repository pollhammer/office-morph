@echo off
:: ============================================================
:: Office-Morph v1.2
:: GitHub: https://github.com/pollhammer/office-morph
:: Author: Manuel Pollhammer
:: ============================================================
setlocal enabledelayedexpansion
title Office-Morph - Manuel Pollhammer

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [!] PLEASE RUN AS ADMINISTRATOR
    pause & exit /b
)

:MENU
cls
echo.
echo    OFFICE-MORPH - v1.2
echo    ------------------------
echo    [1] Start Conversion (Drag-and-Drop or Manual)
echo    [2] Delete Old Files (.doc, .xls, .ppt)
echo    [3] Exit
echo.
set /p "choice=Select an option [1-3]: "

if "%choice%"=="1" goto CONVERT
if "%choice%"=="2" goto DELETE
if "%choice%"=="3" exit
goto MENU

:CONVERT
set "target=%~1"
if "%target%"=="" set /p "target=Please enter path (or press Enter for this folder): "
if "!target!"=="" set "target=%~dp0"
set "target=!target:"=!"
echo.
echo [+] Starting conversion in: "!target!"
powershell.exe -ExecutionPolicy Bypass -File "%~dp0FolderConverter.ps1" -TargetFolder "!target!"
pause
goto MENU

:DELETE
set /p "delpath=Enter path to CLEAN (or press Enter for current folder): "
if "!delpath!"=="" set "delpath=%~dp0"
set "delpath=!delpath:"=!"
echo.
echo [!] WARNING: This will permanently delete all .doc, .xls, and .ppt files in:
echo     "!delpath!"
set /p "confirm=Are you sure? (y/n): "
if /i "!confirm!"=="y" (
    powershell.exe -Command "Get-ChildItem -Path '!delpath!' -Include *.doc, *.xls, *.ppt -Recurse | Remove-Item -Force"
    echo [+] Cleanup complete.
)
pause
goto MENU
