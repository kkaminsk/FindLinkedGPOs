$changePath = Join-Path $PSScriptRoot 'openspec\changes\add-gpo-html-exports'
if (Test-Path $changePath) {
    Remove-Item -Path $changePath -Recurse -Force
    Write-Host "Removed change: add-gpo-html-exports" -ForegroundColor Green
} else {
    Write-Host "Change folder not found" -ForegroundColor Yellow
}
