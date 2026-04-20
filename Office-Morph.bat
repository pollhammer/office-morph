@echo off
:: ============================================================
:: Office-Morph v1.4
:: GitHub: https://github.com
:: Author: Manuel Pollhammer
:: ============================================================
setlocal enabledelayedexpansion
title Office-Morph v1.4 - Manuel Pollhammer

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
echo    OFFICE-MORPH - v1.4 ^| Modernizing Legacy Docs
echo    ----------------------------------------------
echo    [1] %GREEN%Start Conversion%RESET% (Manual Path or Enter for Current)
echo    [2] %RED%Delete Old Files%RESET% (.doc, .xls, .ppt)
echo    [3] Exit
echo.
set /p "choice=Select an option [1-3]: "

if "%choice%"=="1" goto CONVERT
if "%choice%"=="2" goto DELETE
if "%choice%"=="3" exit
goto MENU

:CONVERT
echo.
set "target="
set /p "target=Target Path (Press Enter for current folder): "
if "!target!"=="" set "target=%~dp0%"
set "target=!target:"=!"

:: Check if PS1 exists
if not exist "%~dp0FolderConverter.ps1" (
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
echo.
echo %RED%!!! ATTENTION: This will permanently delete old formats !!!%RESET%
set "delpath="
set /p "delpath=Path to CLEAN (Press Enter for current): "
if "!delpath!"=="" set "delpath=%~dp0%"
set "delpath=!delpath:"=!"

echo %YELLOW%[+] Searching for legacy files in: !delpath!%RESET%
echo.
:: List files first
where /R "!delpath!" *.doc *.xls *.ppt 2>nul
if %errorlevel% neq 0 (
    echo %BLUE%[i] No legacy files found to delete.%RESET%
    pause
    goto MENU
)

set /p "confirm=Are you sure you want to delete these files? [Y/N]: "
if /I "!confirm!"=="Y" (
    del /S /Q "!delpath!\*.doc" "!delpath!\*.xls" "!delpath!\*.ppt"
    echo %GREEN%[+] Cleanup complete.%RESET%
) else (
    echo %BLUE%[i] Cleanup cancelled.%RESET%
)
pause
goto MENU
