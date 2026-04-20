<#
.SYNOPSIS
    OFFICE-MORPH 1.4 - Core Conversion Engine
    
.DESCRIPTION
    Converts legacy .doc, .xls, and .ppt files into modern XML formats (.docx, .xlsx, .pptx).
    
.AUTHOR
    Manuel Pollhammer
    
.YEAR
    2026
#>

param([string]$TargetFolder)

# Fallback to script root if no folder is provided
if ([string]::IsNullOrWhiteSpace($TargetFolder)) { $TargetFolder = $PSScriptRoot }
$TargetFolder = $TargetFolder.Trim('"').Trim("'")
$logFile = Join-Path $TargetFolder "office_morph_details.log"

# Summary Counters
$converted = 0; $skipped = 0; $errors = 0

Write-Host "`n>>> Office-Morph Engine v1.4" -ForegroundColor Cyan
Write-Host ">>> Logging to: $logFile" -ForegroundColor Gray

# Create Log File Header
$header = "====================================================`r`n" +
          "OFFICE-MORPH LOG - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n" +
          "Target Folder: $TargetFolder`r`n" +
          "====================================================`r`n"
$header | Out-File $logFile

$word = $excel = $ppt = $null
function Get-Word  { if (!$script:word)  { $script:word  = New-Object -ComObject Word.Application; $script:word.Visible = $false }; return $script:word }
function Get-Excel { if (!$script:excel) { $script:excel = New-Object -ComObject Excel.Application; $script:excel.DisplayAlerts = $false }; return $script:excel }
function Get-PPT   { if (!$script:ppt)   { $script:ppt   = New-Object -ComObject PowerPoint.Application }; return $script:ppt }

# Search for legacy files (excluding temporary owner files starting with ~$)
$files = Get-ChildItem -Path $TargetFolder -Include *.doc, *.xls, *.ppt -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "~$*" }

if ($null -eq $files) {
    "INFO: No legacy files found." | Out-File $logFile -Append
} else {
    foreach ($file in $files) {
        $newExt = switch ($file.Extension.ToLower()) { ".doc" { ".docx" } ".xls" { ".xlsx" } ".ppt" { ".pptx" } }
        $newPath = [System.IO.Path]::ChangeExtension($file.FullName, $newExt)

        # SKIP LOGIC
        if (Test-Path $newPath) { 
            $skipMsg = "SKIPPED: Target already exists. File: $($file.FullName)"
            Write-Host "Skipped: $($file.Name)" -ForegroundColor Gray
            $skipMsg | Out-File $logFile -Append
            $skipped++; continue 
        }

        # CONVERSION LOGIC
        Write-Host "Morphing: $($file.Name)... " -NoNewline -ForegroundColor Cyan
        try {
            switch ($file.Extension.ToLower()) {
                ".doc" { 
                    $doc = (Get-Word).Documents.Open($file.FullName, $false, $true)
                    $doc.SaveAs2($newPath, 16)
                    $doc.Close() 
                }
                ".xls" { 
                    $wb = (Get-Excel).Workbooks.Open($file.FullName)
                    $wb.SaveAs($newPath, 51)
                    $wb.Close() 
                }
                ".ppt" { 
                    $pres = (Get-PPT).Presentations.Open($file.FullName, 1, 0, 0)
                    $pres.SaveAs($newPath, 24)
                    $pres.Close() 
                }
            }
            Write-Host "SUCCESS" -ForegroundColor Green
            "SUCCESS: Converted $($file.FullName) to $newPath" | Out-File $logFile -Append
            $converted++
        } catch {
            Write-Host "FAILED" -ForegroundColor Red
            
            # Check if the error is related to path length
            $reason = $_.Exception.Message
            if ($_.Exception.InnerException -match "PathTooLongException" -or $file.FullName.Length -gt 255) {
                $reason = "The file path is too long for Windows (limit: 260 characters)."
            }

            $errMsg = "ERROR: Failed to convert $($file.FullName). Reason: $reason"
            $errMsg | Out-File $logFile -Append
            $errors++
        }
    }
}

# Cleanup COM objects
if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
if ($excel) { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
if ($ppt) { $ppt.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null }

# Final Log Summary
$footer = "`r`nSummary: $converted Converted, $skipped Skipped, $errors Errors."
$footer | Out-File $logFile -Append

# Final Console Output
Write-Host "---------------------------------------------------"
Write-Host "Summary: " -NoNewline
Write-Host "$converted Converted  " -ForegroundColor Green -NoNewline
Write-Host "$skipped Skipped  " -ForegroundColor Gray -NoNewline
Write-Host "$errors Errors" -ForegroundColor Red
