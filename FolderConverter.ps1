<#
.SYNOPSIS
    OFFICE-MORPH 1.3 - Core Conversion Engine
    
.DESCRIPTION
    Converts legacy .doc, .xls, and .ppt files into modern XML formats.
    
.AUTHOR
    Manuel Pollhammer
    
.YEAR
    2026
#>

param([string]$TargetFolder)

if ([string]::IsNullOrWhiteSpace($TargetFolder)) { $TargetFolder = $PSScriptRoot }
$TargetFolder = $TargetFolder.Trim('"').Trim("'")

# Summary Counters
$converted = 0; $skipped = 0; $errors = 0

Write-Host ">>> Office-Morph Engine Active" -ForegroundColor Cyan
Write-Host ">>> Target: $TargetFolder" -ForegroundColor Gray
Write-Host "---------------------------------------------------"

$word = $excel = $ppt = $null
function Get-Word  { if (!$script:word)  { $script:word  = New-Object -ComObject Word.Application; $script:word.Visible = $false }; return $script:word }
function Get-Excel { if (!$script:excel) { $script:excel = New-Object -ComObject Excel.Application; $script:excel.DisplayAlerts = $false }; return $script:excel }
function Get-PPT   { if (!$script:ppt)   { $script:ppt   = New-Object -ComObject PowerPoint.Application }; return $script:ppt }

# File Search (Ignoring temp files starting with ~$)
$files = Get-ChildItem -Path $TargetFolder -Include *.doc, *.xls, *.ppt -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "~$*" }

if ($null -eq $files -or $files.Count -eq 0) {
    Write-Host "No convertible files found." -ForegroundColor Gray
} else {
    foreach ($file in $files) {
        $basePath = $file.FullName.Substring(0, $file.FullName.Length - $file.Extension.Length)
        $newExt = switch ($file.Extension.ToLower()) { ".doc" { ".docx" } ".xls" { ".xlsx" } ".ppt" { ".pptx" } }
        $newPath = $basePath + $newExt

        if (Test-Path $newPath) { $skipped++; continue }

        Write-Host "Morphing: $($file.Name)... " -NoNewline -ForegroundColor Cyan
        try {
            switch ($file.Extension.ToLower()) {
                ".doc" { $doc = (Get-Word).Documents.Open($file.FullName); $doc.SaveAs2($newPath, 16); $doc.Close() }
                ".xls" { $wb = (Get-Excel).Workbooks.Open($file.FullName); $wb.SaveAs($newPath, 51); $wb.Close() }
                ".ppt" { $pres = (Get-PPT).Presentations.Open($file.FullName, 0, 0, 0); $pres.SaveAs($newPath, 24); $pres.Close() }
            }
            Write-Host "SUCCESS" -ForegroundColor Green
            $converted++
        } catch {
            Write-Host "FAILED" -ForegroundColor Red
            $errors++
        }
    }
}

# Final Summary Output
Write-Host "---------------------------------------------------"
Write-Host "Summary: " -NoNewline
Write-Host "$converted Converted  " -ForegroundColor Green -NoNewline
Write-Host "$skipped Skipped  " -ForegroundColor Gray -NoNewline
Write-Host "$errors Errors" -ForegroundColor Red

if ($word) { $word.Quit() }
if ($excel) { $excel.Quit() }
if ($ppt) { $ppt.Quit() }
