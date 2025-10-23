$latestFolder = Get-ChildItem "$env:USERPROFILE\Documents" -Directory -Filter "Find-LinkedGPOs-*" | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object -First 1

if ($latestFolder) {
    $logFile = Get-ChildItem $latestFolder.FullName -Filter "*.log" | Select-Object -First 1
    if ($logFile) {
        Write-Host "=== LOG FILE: $($logFile.FullName) ===" -ForegroundColor Cyan
        Write-Host ""
        Get-Content $logFile.FullName
    } else {
        Write-Host "No log file found in folder" -ForegroundColor Red
    }
} else {
    Write-Host "No output folder found" -ForegroundColor Red
}
