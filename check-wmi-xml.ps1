$latestFolder = Get-ChildItem "$env:USERPROFILE\Documents" -Directory -Filter "Find-LinkedGPOs-*" | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object -First 1

if ($latestFolder) {
    $xmlFile = Get-ChildItem $latestFolder.FullName -Filter "*W11Users*.xml" | Select-Object -First 1
    if ($xmlFile) {
        Write-Host "Checking XML file: $($xmlFile.FullName)" -ForegroundColor Cyan
        [xml]$xml = Get-Content $xmlFile.FullName
        $wmiFilters = $xml.SelectNodes("//WmiFilter")
        Write-Host "`nFound $($wmiFilters.Count) WMI filter elements:" -ForegroundColor Yellow
        foreach ($wmi in $wmiFilters) {
            Write-Host "  Name: $($wmi.name)"
            Write-Host "  Id: $($wmi.id)"
            Write-Host "  Query: $($wmi.query)"
            Write-Host "  Has query attribute: $($wmi.HasAttribute('query'))"
            Write-Host ""
        }
    } else {
        Write-Host "No XML file found for W11Users" -ForegroundColor Red
    }
} else {
    Write-Host "No output folder found" -ForegroundColor Red
}
