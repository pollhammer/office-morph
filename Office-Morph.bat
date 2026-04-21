@echo off
:: ============================================================
:: Office-Morph v2.0
:: GitHub: https://github.com/pollhammer/office-morph
:: Author: Manuel Pollhammer
:: ============================================================
mode con: cols=91 lines=24
setlocal enabledelayedexpansion
title Office-Morph v2.0 - Manuel Pollhammer

:: ANSI Colors
for /F "delims=#" %%a in ('"prompt #$E# & for %%b in (1) do rem"') do set "E=%%a"
set "BLUE=%E%[94m"
set "GREEN=%E%[92m"
set "YELLOW=%E%[93m"
set "RED=%E%[91m"
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
echo    OFFICE-MORPH - v2.0 ^| Modernizing Legacy Docs
echo    ----------------------------------------------
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
cls
echo.
echo %BLUE%   ____  ____________________________        %GREEN%   __  _______  ____  ____  __  __
echo %BLUE%  / __ \/ ____/ ____/  _/ ____/ ____/       %GREEN%   /  ^|/  / __ \/ __ \/ __ \/ / / /
echo %BLUE% / / / / /_  / /_   / // /   / __/   %GREEN%______   / /^|_/ / / / / /_/ / /_/ / /_/ / 
echo %BLUE%/ /_/ / __/ / __/ _/ // /___/ /__   %GREEN%/_____/  %GREEN%/ /  / / /_/ / _, _/ ____/ __  /  
echo %BLUE%\____/_/   /_/   /___/\____/_____/       %GREEN%   /_/  /_/\____/_/ ^|_/_/   /_/ /_/ 
echo %RESET%
echo.
set "target="
set /p "target=Target Path (Press Enter for current folder): "
if "!target!"=="" set "target=%~dp0"
set "target=!target:"=!"

if not exist "%~dp0FolderConverter.ps1" (
    echo.
    echo %RED%[!] Error: FolderConverter.ps1 not found in %~dp0%RESET%
    pause
    goto MENU
)

echo.
echo %YELLOW%[+] Initializing Engine...%RESET%
powershell.exe -ExecutionPolicy Bypass -File "%~dp0FolderConverter.ps1" -TargetFolder "!target!"
echo.
echo %GREEN%[+] Process finished.%RESET%
pause
goto MENU

:DELETE
cls
echo.
echo %BLUE%   ____  ____________________________        %GREEN%   __  _______  ____  ____  __  __
echo %BLUE%  / __ \/ ____/ ____/  _/ ____/ ____/       %GREEN%   /  ^|/  / __ \/ __ \/ __ \/ / / /
echo %BLUE% / / / / /_  / /_   / // /   / __/   %GREEN%______   / /^|_/ / / / / /_/ / /_/ / /_/ / 
echo %BLUE%/ /_/ / __/ / __/ _/ // /___/ /__   %GREEN%/_____/  %GREEN%/ /  / / /_/ / _, _/ ____/ __  /  
echo %BLUE%\____/_/   /_/   /___/\____/_____/       %GREEN%   /_/  /_/\____/_/ ^|_/_/   /_/ /_/ 
echo %RESET%
echo.
echo %RED%!!! ATTENTION: This will permanently delete old formats !!!%RESET%
set /p "delpath=Enter path to CLEAN (or press Enter for current folder): "
if "!delpath!"=="" set "delpath=%~dp0"
set "delpath=!delpath:"=!"
echo.
echo     "!delpath!"
set /p "confirm=Are you sure? (y/n): "
if /i "!confirm!"=="y" (
echo.
    powershell.exe -Command "Get-ChildItem -Path '!delpath!' -Include *.doc, *.xls, *.ppt -Recurse | Remove-Item -Force"
    echo %GREEN%[+] Cleanup complete.%RESET%
)
pause
goto MENU


