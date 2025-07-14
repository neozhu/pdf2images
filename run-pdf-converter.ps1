#!/usr/bin/env pwsh
Write-Host "Starting PDF to Image Converter..." -ForegroundColor Green
Write-Host ""
Write-Host "Press Ctrl+C to stop the process gracefully at any time." -ForegroundColor Yellow
Write-Host ""

# Change to script directory
Set-Location $PSScriptRoot

try {
    dotnet run
    Write-Host ""
    Write-Host "PDF conversion completed successfully!" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Press any key to close..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
