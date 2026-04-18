@echo off
:: ============================================================
:: Office-Morph v1.3
:: GitHub: https://github.com/pollhammer/office-morph
:: Author: Manuel Pollhammer
:: ============================================================
setlocal enabledelayedexpansion
title Office-Morph - Manuel Pollhammer

for /F "delims=#" %%a in ('"prompt #$E# & for %%b in (1) do rem"') do set "E=%%a"
set "BLUE=%E%[94m"
set "GREEN=%E%[92m"
set "RESET=%E%[0m"

:MENU
cls
echo.
echo %BLUE%   ____  ____________________________        %GREEN%   __  _______  ____  ____  __  __
echo %BLUE%  / __ \/ ____/ ____/  _/ ____/ ____/       %GREEN%   /  ^|/  / __ \/ __ \/ __ \/ / / /
echo %BLUE% / / / / /_  / /_   / // /   / __/   %GREEN%______   / /^|_/ / / / / /_/ / /_/ / /_/ / 
echo %BLUE%/ /_/ / __/ / __/ _/ // /___/ /__   %GREEN%/_____/  %GREEN%/ /  / / /_/ / _, _/ ____/ __  /  
echo %BLUE%\____/_/   /_/   /___/\____/_____/       %GREEN%   /_/  /_/\____/_/ ^|_/_/   /_/ /_/ 
echo %RESET%
echo    OFFICE-MORPH - v1.3
echo    ------------------------
echo    [1] Start Conversion (Manual Path or Enter for Current)
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
