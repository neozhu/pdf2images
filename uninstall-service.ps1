# PowerShell script to uninstall PDF2Images Windows Service
# Run this script as Administrator

param(
    [string]$ServiceName = "PDF2Images",
    [string]$ServicePath = "C:\Services\PDF2Images",
    [switch]$RemoveFiles = $false
)

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires Administrator privileges. Please run as Administrator." -ForegroundColor Red
    Read-Host "Press any key to exit"
    exit 1
}

Write-Host "Uninstalling PDF2Images Windows Service..." -ForegroundColor Green

try {
    # Check if service exists
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    
    if ($service) {
        Write-Host "Found service: $ServiceName" -ForegroundColor Yellow
        
        # Stop the service if running
        if ($service.Status -eq 'Running') {
            Write-Host "Stopping service..." -ForegroundColor Yellow
            Stop-Service -Name $ServiceName -Force
            Start-Sleep -Seconds 5
        }
        
        # Delete the service
        Write-Host "Removing service..." -ForegroundColor Yellow
        sc.exe delete $ServiceName
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Service removed successfully!" -ForegroundColor Green
        } else {
            Write-Host "Failed to remove service. Error code: $LASTEXITCODE" -ForegroundColor Red
        }
    } else {
        Write-Host "Service '$ServiceName' not found." -ForegroundColor Yellow
    }
    
    # Remove service files if requested
    if ($RemoveFiles -and (Test-Path $ServicePath)) {
        Write-Host "Removing service files from: $ServicePath" -ForegroundColor Yellow
        
        # Confirm before deletion
        $confirmation = Read-Host "Are you sure you want to delete all files in $ServicePath? (y/N)"
        if ($confirmation -eq 'y' -or $confirmation -eq 'Y') {
            Remove-Item -Path $ServicePath -Recurse -Force
            Write-Host "Service files removed." -ForegroundColor Green
        } else {
            Write-Host "Service files preserved." -ForegroundColor Yellow
        }
    }
    
    Write-Host "`nUninstallation Complete!" -ForegroundColor Green
    
    if (-not $RemoveFiles) {
        Write-Host "`nNote: Service files are still located at: $ServicePath" -ForegroundColor Cyan
        Write-Host "Use -RemoveFiles switch to delete them: .\uninstall-service.ps1 -RemoveFiles" -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error occurred during uninstallation: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Read-Host "`nPress any key to exit"
