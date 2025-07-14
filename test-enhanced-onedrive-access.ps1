# Enhanced OneDrive file access test script
# This script simulates the service's OneDrive file handling logic

param(
    [Parameter(Mandatory=$true)]
    [string]$OneDrivePath
)

Write-Host "Enhanced OneDrive file access test..." -ForegroundColor Green
Write-Host "Path: $OneDrivePath" -ForegroundColor Yellow

# Check if path exists
if (-not (Test-Path $OneDrivePath)) {
    Write-Host "ERROR: OneDrive path does not exist: $OneDrivePath" -ForegroundColor Red
    exit 1
}

# Get all PDF files (same logic as in the service)
$pdfFiles = Get-ChildItem -Path $OneDrivePath -Recurse -Filter "*.pdf" | Where-Object {
    $directory = $_.Directory.FullName
    $_.Directory.Name -ne ".pdf" -and 
    -not $directory.Contains("Z1-template") -and
    $_.Name -ne ".pdf" -and
    $_.Length -gt 0
}

Write-Host "Found $($pdfFiles.Count) PDF files to test" -ForegroundColor Green

if ($pdfFiles.Count -eq 0) {
    Write-Host "No PDF files found for testing" -ForegroundColor Yellow
    exit 0
}

# Test with retry logic similar to the service
function Test-FileWithRetry {
    param(
        [System.IO.FileInfo]$File,
        [int]$MaxRetries = 3
    )
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            Write-Host "  Attempt $attempt/$MaxRetries..." -ForegroundColor Cyan
            
            # Check if OneDrive placeholder
            $File.Refresh()
            $isPlaceholder = $File.Attributes -band [System.IO.FileAttributes]::ReparsePoint
            
            if ($isPlaceholder) {
                Write-Host "    OneDrive placeholder detected" -ForegroundColor Yellow
                
                # Try to trigger download
                $stream = [System.IO.File]::OpenRead($File.FullName)
                $bufferSize = [Math]::Min(8192, [Math]::Max($stream.Length, 1024))
                $buffer = New-Object byte[] $bufferSize
                $totalRead = 0
                
                while ($totalRead -lt $bufferSize) {
                    $bytesRead = $stream.Read($buffer, $totalRead, $bufferSize - $totalRead)
                    if ($bytesRead -eq 0) { break }
                    $totalRead += $bytesRead
                }
                $stream.Close()
                
                Write-Host "    Triggered download by reading $totalRead bytes" -ForegroundColor Cyan
                Start-Sleep -Milliseconds 1000
                
                # Check again
                $File.Refresh()
                $isPlaceholder = $File.Attributes -band [System.IO.FileAttributes]::ReparsePoint
            }
            
            # Final verification
            $stream = [System.IO.File]::OpenRead($File.FullName)
            $length = $stream.Length
            
            # Try to read some data to verify accessibility
            if ($length -gt 0) {
                $testBuffer = New-Object byte[] ([Math]::Min(1024, $length))
                $testRead = $stream.Read($testBuffer, 0, $testBuffer.Length)
                Write-Host "    Successfully read $testRead bytes for verification" -ForegroundColor Green
            }
            
            $stream.Close()
            
            $status = if ($isPlaceholder) { "placeholder but accessible" } else { "fully downloaded" }
            Write-Host "    SUCCESS: File is $status, size: $length bytes" -ForegroundColor Green
            return $true
            
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Host "    ERROR on attempt $attempt : $errorMsg" -ForegroundColor Red
            
            if ($attempt -lt $MaxRetries) {
                $delay = $attempt * 500
                Write-Host "    Waiting $delay ms before retry..." -ForegroundColor Yellow
                Start-Sleep -Milliseconds $delay
            }
        }
    }
    
    return $false
}

# Test files
$successCount = 0
$testCount = [Math]::Min(3, $pdfFiles.Count)

Write-Host "`nTesting $testCount files with enhanced retry logic:" -ForegroundColor Green

for ($i = 0; $i -lt $testCount; $i++) {
    $file = $pdfFiles[$i]
    Write-Host "`nFile $($i+1)/$testCount : $($file.Name)" -ForegroundColor Yellow
    Write-Host "  Path: $($file.FullName)" -ForegroundColor Gray
    
    if (Test-FileWithRetry -File $file) {
        $successCount++
    }
}

Write-Host "`n" + "="*50 -ForegroundColor White
Write-Host "ENHANCED TEST SUMMARY:" -ForegroundColor Green
Write-Host "  Total files tested: $testCount" -ForegroundColor Yellow
Write-Host "  Successfully accessed: $successCount" -ForegroundColor Green
Write-Host "  Failed: $($testCount - $successCount)" -ForegroundColor Red

if ($successCount -eq $testCount) {
    Write-Host "`nALL FILES ACCESSIBLE! The enhanced service should work correctly." -ForegroundColor Green
    Write-Host "The service's improved OneDrive handling should resolve the file access issues." -ForegroundColor Green
    exit 0
} else {
    Write-Host "`nSome files still have issues. Check OneDrive sync status and permissions." -ForegroundColor Yellow
    Write-Host "Consider running OneDrive as the same user account that will run the service." -ForegroundColor Yellow
    exit 1
}
