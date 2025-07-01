# PowerShell script to install PDF2Images as Windows Service
# Run this script as Administrator

param(
    [string]$ServicePath = "C:\Services\PDF2Images",
    [string]$ServiceName = "PDF2Images",
    [string]$DisplayName = "PDF to Images Converter Service",
    [string]$Description = "Automatically converts PDF files to images from OneDrive directory"
)

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires Administrator privileges. Please run as Administrator." -ForegroundColor Red
    Read-Host "Press any key to exit"
    exit 1
}

Write-Host "Installing PDF2Images Windows Service..." -ForegroundColor Green

try {
    # Create service directory if it doesn't exist
    if (-not (Test-Path $ServicePath)) {
        Write-Host "Creating service directory: $ServicePath" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $ServicePath -Force | Out-Null
    }

    # Copy published files to service directory
    $publishPath = Join-Path $PSScriptRoot "publish"
    if (Test-Path $publishPath) {
        Write-Host "Copying application files to service directory..." -ForegroundColor Yellow
        Copy-Item -Path "$publishPath\*" -Destination $ServicePath -Recurse -Force
    } else {
        Write-Host "Published files not found. Please run 'dotnet publish -c Release -o publish' first." -ForegroundColor Red
        exit 1
    }

    # Stop service if it exists
    $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($existingService) {
        Write-Host "Stopping existing service..." -ForegroundColor Yellow
        Stop-Service -Name $ServiceName -Force
        Write-Host "Removing existing service..." -ForegroundColor Yellow
        sc.exe delete $ServiceName
        Start-Sleep -Seconds 2
    }

    # Create the service
    $exePath = Join-Path $ServicePath "pdf2images.exe"
    Write-Host "Creating Windows Service..." -ForegroundColor Yellow
    
    $serviceCommand = "sc.exe create `"$ServiceName`" binpath= `"$exePath`" displayname= `"$DisplayName`" description= `"$Description`" start= auto"
    Invoke-Expression $serviceCommand

    if ($LASTEXITCODE -eq 0) {
        Write-Host "Service created successfully!" -ForegroundColor Green
        
        # Set service to restart on failure
        Write-Host "Configuring service recovery options..." -ForegroundColor Yellow
        sc.exe failure $ServiceName reset= 86400 actions= restart/5000/restart/5000/restart/5000
        
        # Start the service
        Write-Host "Starting service..." -ForegroundColor Yellow
        Start-Service -Name $ServiceName
        
        # Check service status
        $service = Get-Service -Name $ServiceName
        Write-Host "Service Status: $($service.Status)" -ForegroundColor $(if ($service.Status -eq 'Running') { 'Green' } else { 'Red' })
        
        Write-Host "`nService Installation Complete!" -ForegroundColor Green
        Write-Host "Service Name: $ServiceName" -ForegroundColor Cyan
        Write-Host "Service Path: $ServicePath" -ForegroundColor Cyan
        Write-Host "Log Files: $ServicePath\logs\" -ForegroundColor Cyan
        Write-Host "`nYou can manage the service using:" -ForegroundColor Yellow
        Write-Host "- Services.msc (GUI)" -ForegroundColor White
        Write-Host "- net start $ServiceName" -ForegroundColor White
        Write-Host "- net stop $ServiceName" -ForegroundColor White
        Write-Host "- Get-Service $ServiceName" -ForegroundColor White
    } else {
        Write-Host "Failed to create service. Error code: $LASTEXITCODE" -ForegroundColor Red
        exit 1
    }
}
catch {
    Write-Host "Error occurred during installation: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Read-Host "`nPress any key to exit"
