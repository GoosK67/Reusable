#!/usr/bin/env powershell
"""
Monitor script to track sd_chapter_classifier.py progress
"""

$logFile = "c:\Users\koengo\OneDrive - Cegeka\Documents\Reusable Assets\Products\sd_classification_full.log"
$outputFile = "c:\Users\koengo\OneDrive - Cegeka\Documents\Reusable Assets\Products\sd_chapter_classification.xlsx"

Write-Host "=== Chapter Classifier Progress Monitor ===" -ForegroundColor Cyan
Write-Host "Log file: $logFile`n" -ForegroundColor Gray

$previousCount = 0
while ($true) {
    $logExists = Test-Path $logFile
    
    if ($logExists) {
        $content = Get-Content $logFile -ErrorAction SilentlyContinue
        
        # Count processed files
        $matches = $content | Select-String "\[\d+/95\]" -AllMatches
        $processedCount = $matches.Count
        
        # Show progress
        if ($processedCount -gt $previousCount) {
            $lastLine = ($content | Select-String "\[\d+/95\]" | Select-Object -Last 1)
            Write-Host "$(Get-Date -Format 'HH:mm:ss') | Processed: $processedCount/95 | $lastLine" -ForegroundColor Green
            $previousCount = $processedCount
        }
        
        # Check if Python process still running
        $pyRunning = Get-Process python -ErrorAction SilentlyContinue
        if (-not $pyRunning) {
            Write-Host "`n✅ COMPLETED!" -ForegroundColor Green
            Write-Host "Script finished at $(Get-Date -Format 'HH:mm:ss')"
            
            if (Test-Path $outputFile) {
                $size = (Get-Item $outputFile).Length / 1MB
                Write-Host "Output file: $size MB" -ForegroundColor Green
            }
            break
        }
    }
    
    Start-Sleep -Seconds 30
}
