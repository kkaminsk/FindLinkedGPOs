$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
Write-Host "PATH reloaded successfully!" -ForegroundColor Green
Write-Host "You can now use: node, npm, and openspec" -ForegroundColor Cyan
