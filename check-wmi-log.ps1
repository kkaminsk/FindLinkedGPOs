$latestLog = Get-ChildItem "$env:USERPROFILE\Documents" -Filter "Find-LinkedGPOs-*.log" | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object -First 1

if ($latestLog) {
    Write-Host "Checking log: $($latestLog.FullName)" -ForegroundColor Cyan
    Write-Host "`n=== WMI-Related Log Entries ===" -ForegroundColor Yellow
    Get-Content $latestLog.FullName | Select-String -Pattern "WMI|Extracting|cached|query" -CaseSensitive:$false
} else {
    Write-Host "No log file found" -ForegroundColor Red
}
