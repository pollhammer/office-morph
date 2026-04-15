@echo off
:: #######################################################
:: # PROJECT: OFFICE MORPH 2026 - Version 1.0
:: # AUTHOR:  Manuel Pollhammer
:: # DATE:    2026
:: # INFO:    Requires Local Admin Rights
:: #######################################################

setlocal enabledelayedexpansion
title Office Morph 2026 - Manuel Pollhammer

:: Administrator-Check
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [!] BITTE ALS ADMINISTRATOR AUSFUEHREN
    echo Rechtsklick auf die Datei -> Als Administrator ausfuehren
    pause
    exit /b
)

:: 1. Check: Drag-and-Drop
set "target=%~1"

if "%target%"=="" (
    echo.
    echo    OFFICE MORPH 2026
    echo    -----------------
    echo.
    set /p "target=Bitte Pfad eingeben (oder Enter fuer diesen Ordner): "
)

:: 2. Check: Default auf aktuellen Ordner
if "!target!"=="" (
    set "target=%~dp0"
)

:: Bereinigung
set "target=!target:"=!"

echo.
echo [+] Starte Konvertierung in: "!target!"
echo.

powershell.exe -ExecutionPolicy Bypass -File "%~dp0FolderConverter.ps1" -TargetFolder "!target!"

echo.
echo [+] Vorgang abgeschlossen.
pause
