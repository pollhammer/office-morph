<#
.SYNOPSIS
    OFFICE MORPH 2026 - Core Conversion Engine
    
.DESCRIPTION
    Konvertiert .doc, .xls und .ppt in moderne Formate.
    
.AUTHOR
    Manuel Pollhammer
    
.YEAR
    2026
#>

param([string]$TargetFolder)

# Initialisierung & Pfad-Fix
if ([string]::IsNullOrWhiteSpace($TargetFolder)) { $TargetFolder = $PSScriptRoot }
$TargetFolder = $TargetFolder.Trim('"').Trim("'")

if (-not (Test-Path $TargetFolder)) {
    Write-Host "[!] Pfad nicht gefunden: $TargetFolder" -ForegroundColor Red
    return
}

Write-Host ">>> Office Morph Engine aktiv" -ForegroundColor Cyan
Write-Host ">>> Ziel: $TargetFolder" -ForegroundColor Gray
Write-Host "---------------------------------------------------"

$word = $excel = $ppt = $null
function Get-Word  { if (!$script:word)  { $script:word  = New-Object -ComObject Word.Application; $script:word.Visible = $false }; return $script:word }
function Get-Excel { if (!$script:excel) { $script:excel = New-Object -ComObject Excel.Application; $script:excel.DisplayAlerts = $false }; return $script:excel }
function Get-PPT   { if (!$script:ppt)   { $script:ppt   = New-Object -ComObject PowerPoint.Application }; return $script:ppt }

# Dateisuche
$files = Get-ChildItem -Path $TargetFolder -Include *.doc, *.xls, *.ppt -Recurse -ErrorAction SilentlyContinue | Select-Object -Unique

if ($null -eq $files -or $files.Count -eq 0) {
    Write-Host "Keine konvertierbaren Dateien gefunden." -ForegroundColor Gray
} else {
    foreach ($file in $files) {
        $basePath = $file.FullName.Substring(0, $file.FullName.Length - $file.Extension.Length)
        $newExt = switch ($file.Extension.ToLower()) { ".doc" { ".docx" } ".xls" { ".xlsx" } ".ppt" { ".pptx" } }
        $newPath = $basePath + $newExt

        if (Test-Path $newPath) { continue }

        Write-Host "Morphe: $($file.Name)... " -NoNewline -ForegroundColor Cyan
        
        try {
            switch ($file.Extension.ToLower()) {
                ".doc" { 
                    $doc = (Get-Word).Documents.Open($file.FullName)
                    $doc.SaveAs2($newPath, 16)
                    $doc.Close() 
                }
                ".xls" { 
                    $wb = (Get-Excel).Workbooks.Open($file.FullName)
                    $wb.SaveAs($newPath, 51)
                    $wb.Close() 
                }
                ".ppt" { 
                    $pres = (Get-PPT).Presentations.Open($file.FullName, 0, 0, 0)
                    $pres.SaveAs($newPath, 24)
                    $pres.Close() 
                }
            }
            Write-Host "ERFOLG" -ForegroundColor Green
        } catch {
            Write-Host "FEHLER" -ForegroundColor Red
        }
    }
}

# Cleanup
if ($word)  { $word.Quit() }
if ($excel) { $excel.Quit() }
if ($ppt)   { $ppt.Quit() }
