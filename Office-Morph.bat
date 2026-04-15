@echo off
:: ============================================================
:: Office-Morph v1.1
:: GitHub: https://github.com/pollhammer/office-morph
:: Author: Manuel Pollhammer
:: ============================================================

setlocal enabledelayedexpansion
title Office-Morph v1.1 - Manuel Pollhammer

:: Administrator Check
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [!] PLEASE RUN AS ADMINISTRATOR
    echo Right-click the file -^> Run as administrator
    pause
    exit /b
)

:: 1. Check: Drag-and-Drop
set "target=%~1"

if "%target%"=="" (
    echo.
    echo    OFFICE-MORPH v1.1
    echo    -----------------
    echo.
    set /p "target=Please enter path (or press Enter for this folder): "
)

:: 2. Check: Default to current folder
if "!target!"=="" (
    set "target=%~dp0"
)

:: Cleanup
set "target=!target:"=!"

echo.
echo [+] Starting conversion in: "!target!"
echo.

powershell.exe -ExecutionPolicy Bypass -File "%~dp0FolderConverter.ps1" -TargetFolder "!target!"

echo.
echo [+] Process completed.
pause
