@echo off
echo Starting PDF to Image Converter...
echo.
echo Press Ctrl+C to stop the process gracefully at any time.
echo.
cd /d "%~dp0"
dotnet run
echo.
echo PDF conversion completed.
pause
