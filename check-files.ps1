$latestFolder = Get-ChildItem "$env:USERPROFILE\Documents" -Filter "Find-LinkedGPOs-*" -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestFolder) {
    Write-Host "Checking folder: $($latestFolder.FullName)"
    Write-Host "`nLinked GPO files:"
    Get-ChildItem "$($latestFolder.FullName)\GPO\linked" | Select-Object Name, Extension | Format-Table
    Write-Host "`nUnlinked GPO files:"
    Get-ChildItem "$($latestFolder.FullName)\GPO\unlinked" | Select-Object Name, Extension | Format-Table
}
