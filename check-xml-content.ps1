$latestFolder = Get-ChildItem "$env:USERPROFILE\Documents" -Filter "Find-LinkedGPOs-*" -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestFolder) {
    $xmlFile = Join-Path $latestFolder.FullName "GPO\linked\BlockCMD.xml"
    Write-Host "Checking XML file: $xmlFile"
    Write-Host "`nFirst 10 lines:"
    Get-Content $xmlFile -First 10
}
