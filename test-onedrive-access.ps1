# Test OneDrive file access script
# This script helps verify OneDrive file access before running the service

param(
    [Parameter(Mandatory=$true)]
    [string]$OneDrivePath
)

Write-Host "Testing OneDrive file access..." -ForegroundColor Green
Write-Host "Path: $OneDrivePath" -ForegroundColor Yellow

# Check if path exists
if (-not (Test-Path $OneDrivePath)) {
    Write-Host "ERROR: OneDrive path does not exist: $OneDrivePath" -ForegroundColor Red
    exit 1
}

# Get all PDF files
$pdfFiles = Get-ChildItem -Path $OneDrivePath -Recurse -Filter "*.pdf" | Where-Object {
    $_.Directory.Name -ne ".pdf" -and 
    $_.Directory.FullName -notlike "*Z1-template*" -and
    $_.Name -ne ".pdf" -and
    $_.Length -gt 0
}

Write-Host "Found $($pdfFiles.Count) PDF files" -ForegroundColor Green

# Test access to first few files
$testCount = [Math]::Min(5, $pdfFiles.Count)
$successCount = 0
$onlineOnlyCount = 0

for ($i = 0; $i -lt $testCount; $i++) {
    $file = $pdfFiles[$i]
    Write-Host "Testing file $($i+1)/$testCount : $($file.Name)" -ForegroundColor Yellow
    
    try {
        # Check if file is online-only (reparse point)
        if ($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            Write-Host "  - OneDrive online-only file detected" -ForegroundColor Cyan
            $onlineOnlyCount++
            
            # Try to trigger download
            Write-Host "  - Attempting to trigger download..." -ForegroundColor Cyan
            $stream = [System.IO.File]::OpenRead($file.FullName)
            $buffer = New-Object byte[] 4096
            $stream.Read($buffer, 0, $buffer.Length) | Out-Null
            $stream.Close()
            
            # Check again
            $file.Refresh()
            if ($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                Write-Host "  - Still online-only after read attempt" -ForegroundColor Yellow
            } else {
                Write-Host "  - Successfully downloaded" -ForegroundColor Green
            }
        } else {
            Write-Host "  - File is already downloaded locally" -ForegroundColor Green
        }
        
        # Test file access
        $fs = [System.IO.File]::OpenRead($file.FullName)
        $length = $fs.Length
        $fs.Close()
        Write-Host "  - File accessible, size: $length bytes" -ForegroundColor Green
        $successCount++
        
    } catch {
        Write-Host "  - ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nSummary:" -ForegroundColor Green
Write-Host "  Total files tested: $testCount" -ForegroundColor Yellow
Write-Host "  Successfully accessed: $successCount" -ForegroundColor Green
Write-Host "  Online-only files found: $onlineOnlyCount" -ForegroundColor Cyan
Write-Host "  Failed: $($testCount - $successCount)" -ForegroundColor Red

if ($successCount -eq $testCount) {
    Write-Host "`nAll files are accessible! The service should work correctly." -ForegroundColor Green
    exit 0
} else {
    Write-Host "`nSome files are not accessible. Check OneDrive sync status." -ForegroundColor Yellow
    exit 1
}
