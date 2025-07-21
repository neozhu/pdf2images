using PDFtoImage;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.FileProperties;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace pdf2images
{
    public class Pdf2ImageService
    {
        private readonly string _sourcePath;
        private readonly EmailService _emailService;
        private readonly ILogger<Pdf2ImageService> _logger;

        public Pdf2ImageService(IConfiguration configuration, EmailService emailService, ILogger<Pdf2ImageService> logger)
        {
            // Read the OneDrive:Path configuration
            _sourcePath = configuration["OneDrive:Path"]
                ?? throw new ArgumentException("Please set OneDrive:Path in the configuration file");
            _emailService = emailService;
            _logger = logger;
        }

        /// <summary>
        /// Processes PDF files in smaller batches to prevent excessive memory usage
        /// </summary>
        public async Task ProcessAllAsync(int batchSize = 10, CancellationToken cancellationToken = default)
        {
            _logger.LogInformation("Starting PDF conversion task");
            
            int totalPdfCount = 0;
            int successfulPdfCount = 0;
            int failedPdfCount = 0;
            int totalPageCount = 0;
            var errorList = new List<string>();
            var startTime = DateTime.Now;
            
            // Dictionary to track statistics by directory
            var directoryStats = new Dictionary<string, (int TotalFiles, int SuccessFiles, int FailedFiles, int TotalPages)>();

            try
            {
                var pdfFiles = Directory
                    .EnumerateFiles(_sourcePath, "*.pdf", SearchOption.AllDirectories)
                    .Where(p => {
                        var directory = Path.GetDirectoryName(p);
                        return directory != null && 
                               !directory.EndsWith(Path.DirectorySeparatorChar + ".pdf") && 
                               !directory.Contains("Z1-template") &&
                               !directory.EndsWith(Path.DirectorySeparatorChar + ".broken");
                    })
                    .ToList();

                totalPdfCount = pdfFiles.Count;
                _logger.LogInformation($"Found {totalPdfCount} PDF files to process");

                // Process files in batches to limit memory usage
                for (int i = 0; i < pdfFiles.Count; i += batchSize)
                {
                    // Check for cancellation between batches
                    cancellationToken.ThrowIfCancellationRequested();
                    
                    // Get the current batch
                    var batch = pdfFiles.Skip(i).Take(batchSize).ToList();
                    _logger.LogInformation($"Processing batch {i / batchSize + 1} of {(pdfFiles.Count + batchSize - 1) / batchSize} ({batch.Count} files)");
                    
                    foreach (var pdfPath in batch)
                    {
                        // Check for cancellation between files
                        cancellationToken.ThrowIfCancellationRequested();
                        
                        var fileBase = Path.GetFileNameWithoutExtension(pdfPath);
                        var pdfDirectory = Path.GetDirectoryName(pdfPath);
                        
                        // Skip if directory is null (shouldn't happen, but for safety)
                        if (string.IsNullOrEmpty(pdfDirectory))
                        {
                            _logger.LogError($"Could not determine directory for file: {pdfPath}");
                            continue;
                        }
                        
                        // Create a .pdf archive folder in the PDF file's directory
                        var archiveDirectory = Path.Combine(pdfDirectory, ".pdf");
                        Directory.CreateDirectory(archiveDirectory);

                        //Create a .archive folder in the PDF file's directory
                        var archivePath = Path.Combine(pdfDirectory, ".archive");
                        if(!Directory.Exists(archivePath))
                        {
                            Directory.CreateDirectory(archivePath);
                        }

                        //Create a .broken folder in the PDF file's directory for corrupted/invalid files
                        var brokenPath = Path.Combine(pdfDirectory, ".broken");
                        if(!Directory.Exists(brokenPath))
                        {
                            Directory.CreateDirectory(brokenPath);
                        }

                        // Initialize directory stats if not already present
                        if (!directoryStats.ContainsKey(pdfDirectory))
                        {
                            directoryStats[pdfDirectory] = (0, 0, 0, 0);
                        }
                        
                        // Update the total files count for this directory
                        var currentStats = directoryStats[pdfDirectory];
                        directoryStats[pdfDirectory] = (currentStats.TotalFiles + 1, 
                                                       currentStats.SuccessFiles, 
                                                       currentStats.FailedFiles, 
                                                       currentStats.TotalPages);

                        try
                        {
                            _logger.LogInformation($"Started processing: {pdfPath}");
                            
                            // Use a timeout for individual file processing
                            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
                            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, cancellationToken);
                            
                            // Process this single file with timeout
                            await ProcessSingleFileAsync(pdfPath, pdfDirectory, fileBase, directoryStats, linkedCts.Token);
                            
                            successfulPdfCount++;
                        }
                        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                        {
                            _logger.LogWarning($"Processing of {pdfPath} was canceled due to service cancellation request");
                            throw; // Re-throw to stop the entire process
                        }
                        catch (OperationCanceledException)
                        {
                            // File processing took too long
                            failedPdfCount++;
                            string errorMessage = $"Processing of {pdfPath} timed out after 5 minutes";
                            _logger.LogError(errorMessage);
                            errorList.Add(errorMessage);
                            
                            // Update directory stats for failed conversion
                            currentStats = directoryStats[pdfDirectory];
                            directoryStats[pdfDirectory] = (currentStats.TotalFiles, 
                                                           currentStats.SuccessFiles, 
                                                           currentStats.FailedFiles + 1, 
                                                           currentStats.TotalPages);
                        }
                        catch (Exception ex)
                        {
                            failedPdfCount++;
                            string errorMessage = $"Error processing file {pdfPath}: {ex.Message}";
                            _logger.LogError(ex, errorMessage);
                            errorList.Add(errorMessage);
                            
                            // Update directory stats for failed conversion
                            currentStats = directoryStats[pdfDirectory];
                            directoryStats[pdfDirectory] = (currentStats.TotalFiles, 
                                                           currentStats.SuccessFiles, 
                                                           currentStats.FailedFiles + 1, 
                                                           currentStats.TotalPages);
                            
                            // Send error notification email
                            try
                            {
                                await SendErrorNotificationAsync(pdfPath, ex);
                            }
                            catch (Exception emailEx)
                            {
                                _logger.LogError(emailEx, "Failed to send error notification email");
                            }
                        }
                        
                        // Force garbage collection between processing large files
                        if (i % 5 == 0)
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }
                    
                    // Wait a bit between batches to let system resources recover
                    try
                    {
                        await Task.Delay(TimeSpan.FromSeconds(10), cancellationToken);
                    }
                    catch (OperationCanceledException)
                    {
                        // Ignore if we're shutting down
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.LogWarning("PDF processing was canceled");
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error during PDF batch processing");
                throw;
            }
            finally
            {
                // Always try to send the summary email, even if processing was interrupted
                try
                {
                    await SendCompletionSummaryAsync(totalPdfCount, successfulPdfCount, failedPdfCount, totalPageCount, errorList, startTime, directoryStats);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to send completion summary email");
                }
            }
        }

        // Helper method to process a single file with cancellation support
        private async Task ProcessSingleFileAsync(string pdfPath, string pdfDirectory, string fileBase, 
            Dictionary<string, (int TotalFiles, int SuccessFiles, int FailedFiles, int TotalPages)> directoryStats,
            CancellationToken cancellationToken)
        {
            try
            {
                // Ensure OneDrive file is downloaded before processing
                await EnsureFileIsDownloadedAsync(pdfPath, cancellationToken);
                
                // Get the number of pages with limited memory usage and retry logic
                int pageCount = 0;
                const int maxReadRetries = 3;
                
                for (int attempt = 0; attempt < maxReadRetries; attempt++)
                {
                    try
                    {
                        // 验证文件是否为有效的PDF格式
                        if (!IsValidPdfFile(pdfPath))
                        {
                            throw new InvalidDataException($"File is not a valid PDF format: {pdfPath}");
                        }
                        
                        using (var pdfStream = File.OpenRead(pdfPath))
                        {
                            pageCount = Conversion.GetPageCount(pdfStream);
                        }
                        break; // Success, exit retry loop
                    }
                    catch (UnauthorizedAccessException ex) when (attempt < maxReadRetries - 1)
                    {
                        _logger.LogWarning($"Access denied reading PDF (attempt {attempt + 1}), retrying: {pdfPath}. Error: {ex.Message}");
                        await Task.Delay(1000 * (attempt + 1), cancellationToken);
                        continue;
                    }
                    catch (IOException ex) when (attempt < maxReadRetries - 1)
                    {
                        _logger.LogWarning($"IO error reading PDF (attempt {attempt + 1}), retrying: {pdfPath}. Error: {ex.Message}");
                        await Task.Delay(1000 * (attempt + 1), cancellationToken);
                        continue;
                    }
                    catch (InvalidDataException ex) when (attempt < maxReadRetries - 1)
                    {
                        _logger.LogWarning($"PDF format validation failed (attempt {attempt + 1}), retrying: {pdfPath}. Error: {ex.Message}");
                        // 等待更长时间，可能是OneDrive文件还在下载
                        await Task.Delay(2000 * (attempt + 1), cancellationToken);
                        continue;
                    }
                    catch (Exception ex) when (ex.Message.Contains("File not in PDF format") && attempt < maxReadRetries - 1)
                    {
                        _logger.LogWarning($"PDF format error (attempt {attempt + 1}), retrying: {pdfPath}. Error: {ex.Message}");
                        // 尝试重新触发OneDrive下载
                        await EnsureFileIsDownloadedAsync(pdfPath, cancellationToken);
                        await Task.Delay(2000 * (attempt + 1), cancellationToken);
                        continue;
                    }
                    catch (InvalidDataException ex) when (attempt == maxReadRetries - 1)
                    {
                        // 最后一次尝试，移动到.broken目录
                        string reason = $"Invalid PDF format after {maxReadRetries} attempts: {ex.Message}";
                        await MoveBrokenFileAsync(pdfPath, reason);
                        throw new InvalidDataException($"File moved to .broken directory. {reason}");
                    }
                    catch (Exception ex) when (ex.Message.Contains("File not in PDF format") && attempt == maxReadRetries - 1)
                    {
                        // 最后一次尝试，移动到.broken目录
                        string reason = $"PDF format error after {maxReadRetries} attempts: {ex.Message}";
                        await MoveBrokenFileAsync(pdfPath, reason);
                        throw new InvalidDataException($"File moved to .broken directory. {reason}");
                    }
                }
                
                // Ensure we actually got the page count
                if (pageCount == 0)
                {
                    string errorMessage = $"Failed to read page count from PDF after {maxReadRetries} attempts: {pdfPath}";
                    await MoveBrokenFileAsync(pdfPath, $"Zero page count after {maxReadRetries} attempts");
                    throw new IOException(errorMessage);
                }
                
                int totalPageCount = 0;
                
                // Process each page individually to save memory
                for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
                {
                    // Check for cancellation between pages
                    cancellationToken.ThrowIfCancellationRequested();
                    
                    // Use the original PDF directory as the image save path
                    string imageName = $"{fileBase}-{pageIndex + 1}.jpeg";
                    string imagePath = Path.Combine(pdfDirectory, imageName);

                    // Open streams for only the current page with retry logic
                    bool pageProcessed = false;
                    for (int pageAttempt = 0; pageAttempt < 2; pageAttempt++)
                    {
                        try
                        {
                            using (var inStream = File.OpenRead(pdfPath))
                            using (var outStream = File.Create(imagePath))
                            {
                                Conversion.SaveJpeg(
                                    outStream,
                                    inStream,
                                    pageIndex,
                                    options: new(
                                        Dpi: 150,
                                        Height: 1200,
                                        WithAnnotations: true,
                                        WithFormFill: true,
                                        WithAspectRatio: true
                                    )
                                );
                            }
                            pageProcessed = true;
                            break; // Success, exit retry loop
                        }
                        catch (Exception ex) when (pageAttempt < 1)
                        {
                            // 检查是否是PDF格式错误
                            if (ex.Message.Contains("File not in PDF format") || ex.Message.Contains("corrupted"))
                            {
                                _logger.LogWarning($"PDF format error processing page {pageIndex + 1} (attempt {pageAttempt + 1}): {ex.Message}");
                                // 如果是格式错误，尝试重新确保文件下载
                                await EnsureFileIsDownloadedAsync(pdfPath, cancellationToken);
                            }
                            else
                            {
                                _logger.LogWarning($"Error processing page {pageIndex + 1} (attempt {pageAttempt + 1}), retrying: {ex.Message}");
                            }
                            await Task.Delay(500, cancellationToken);
                        }
                        catch (Exception ex) when (pageAttempt == 1 && (ex.Message.Contains("File not in PDF format") || ex.Message.Contains("corrupted")))
                        {
                            // 最后一次尝试仍然是PDF格式错误，移动到.broken目录
                            string reason = $"PDF format error during page processing: {ex.Message}";
                            await MoveBrokenFileAsync(pdfPath, reason);
                            throw new InvalidDataException($"File moved to .broken directory. {reason}");
                        }
                    }
                    
                    if (!pageProcessed)
                    {
                        throw new IOException($"Failed to process page {pageIndex + 1} after retries");
                    }
                    
                    totalPageCount++;
                    
                    // Brief pause between pages to prevent CPU spikes
                    if (pageIndex % 5 == 0 && pageIndex > 0)
                    {
                        try
                        {
                            await Task.Delay(TimeSpan.FromMilliseconds(100), cancellationToken);
                        }
                        catch (OperationCanceledException)
                        {
                            // Ignore if we're shutting down
                        }
                    }
                }

                _logger.LogInformation($"Successfully processed {pdfPath}, total {pageCount} pages");

                // After conversion, archive the original PDF to the .pdf subfolder in the PDF's directory
                MoveToArchive(pdfPath, pdfDirectory);
                
                // Update directory stats for successful conversion
                var currentStats = directoryStats[pdfDirectory];
                directoryStats[pdfDirectory] = (currentStats.TotalFiles, 
                                               currentStats.SuccessFiles + 1, 
                                               currentStats.FailedFiles, 
                                               currentStats.TotalPages + pageCount);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error in ProcessSingleFileAsync for file: {pdfPath}");
                
                // Add more specific error information for OneDrive-related issues
                if (ex is UnauthorizedAccessException)
                {
                    _logger.LogError($"Access denied to file - this may be a OneDrive permissions issue: {pdfPath}");
                }
                else if (ex is IOException ioEx)
                {
                    _logger.LogError($"IO error accessing file - this may be a OneDrive sync issue: {pdfPath}. Details: {ioEx.Message}");
                }
                
                throw; // Re-throw to be handled by the calling method
            }
        }

        /// <summary>
        /// Sends a summary email after processing completion.
        /// Only sends email if at least one PDF was successfully converted.
        /// </summary>
        private async Task SendCompletionSummaryAsync(int totalPdfCount, int successfulPdfCount, int failedPdfCount, 
            int totalPageCount, List<string> errorList, DateTime startTime, 
            Dictionary<string, (int TotalFiles, int SuccessFiles, int FailedFiles, int TotalPages)> directoryStats)
        {
            // Skip sending email if no PDFs were successfully converted
            if (successfulPdfCount <= 0)
            {
                _logger.LogInformation("No PDFs were successfully converted. Skipping summary email.");
                return;
            }

            var duration = DateTime.Now - startTime;
            
            StringBuilder bodyBuilder = new StringBuilder();
            bodyBuilder.AppendLine("<h2>PDF to Image Conversion Summary</h2>");
            bodyBuilder.AppendLine("<div style='font-family: Arial, sans-serif; padding: 15px;'>");
            
            // Overall summary table
            bodyBuilder.AppendLine("<h3>Overall Summary</h3>");
            bodyBuilder.AppendLine("<table style='border-collapse: collapse; width: 100%; margin-bottom: 20px;'>");
            bodyBuilder.AppendLine("<tr style='background-color: #f2f2f2;'><th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Item</th><th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Value</th></tr>");
            
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Processing Start Time</td><td style='padding: 10px; border: 1px solid #ddd;'>{startTime:yyyy-MM-dd HH:mm:ss}</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Processing End Time</td><td style='padding: 10px; border: 1px solid #ddd;'>{DateTime.Now:yyyy-MM-dd HH:mm:ss}</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Processing Duration</td><td style='padding: 10px; border: 1px solid #ddd;'>{duration.Hours} hr {duration.Minutes} min {duration.Seconds} sec</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Total PDF Files</td><td style='padding: 10px; border: 1px solid #ddd;'>{totalPdfCount}</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Successfully Processed Files</td><td style='padding: 10px; border: 1px solid #ddd;'>{successfulPdfCount}</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Failed Files</td><td style='padding: 10px; border: 1px solid #ddd; color: {(failedPdfCount > 0 ? "red" : "inherit")};'>{failedPdfCount}</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Total Converted Pages</td><td style='padding: 10px; border: 1px solid #ddd;'>{totalPageCount}</td></tr>");
            bodyBuilder.AppendLine("</table>");
            
            // Directory-based summary
            bodyBuilder.AppendLine("<h3>Directory Summary</h3>");
            bodyBuilder.AppendLine("<table style='border-collapse: collapse; width: 100%;'>");
            bodyBuilder.AppendLine("<tr style='background-color: #f2f2f2;'>" +
                "<th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Directory</th>" +
                "<th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Total Files</th>" +
                "<th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Successful</th>" +
                "<th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Failed</th>" +
                "<th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Pages Converted</th></tr>");
            
            // Sort directories by relative path for better readability
            foreach (var dirStats in directoryStats.OrderBy(d => GetRelativePath(d.Key)))
            {
                string relativeDir = GetRelativePath(dirStats.Key);
                bodyBuilder.AppendLine(
                    $"<tr>" +
                    $"<td style='padding: 10px; border: 1px solid #ddd;'>{relativeDir}</td>" +
                    $"<td style='padding: 10px; border: 1px solid #ddd;'>{dirStats.Value.TotalFiles}</td>" +
                    $"<td style='padding: 10px; border: 1px solid #ddd;'>{dirStats.Value.SuccessFiles}</td>" +
                    $"<td style='padding: 10px; border: 1px solid #ddd; color: {(dirStats.Value.FailedFiles > 0 ? "red" : "inherit")};'>{dirStats.Value.FailedFiles}</td>" +
                    $"<td style='padding: 10px; border: 1px solid #ddd;'>{dirStats.Value.TotalPages}</td>" +
                    $"</tr>");
            }
            
            bodyBuilder.AppendLine("</table>");

            // If there are errors, add the error list below
            if (errorList.Count > 0)
            {
                bodyBuilder.AppendLine("<h3 style='color: red; margin-top: 20px;'>List of Errors:</h3>");
                bodyBuilder.AppendLine("<ul style='color: #555;'>");
                foreach (var error in errorList)
                {
                    bodyBuilder.AppendLine($"<li>{error}</li>");
                }
                bodyBuilder.AppendLine("</ul>");
            }
            
            bodyBuilder.AppendLine("</div>");

            string subject = $"PDF Conversion Completed - Success:{successfulPdfCount} Failure:{failedPdfCount} Total Pages:{totalPageCount}";
            await _emailService.SendEmailAsync(subject, bodyBuilder.ToString(), true);
            _logger.LogInformation("Summary email sent");
        }

        /// <summary>
        /// Gets the relative path from the _sourcePath for better readability.
        /// </summary>
        private string GetRelativePath(string fullPath)
        {
            // If the path is contained within the source path, get the relative portion
            if (fullPath.StartsWith(_sourcePath, StringComparison.OrdinalIgnoreCase))
            {
                string relativePath = fullPath.Substring(_sourcePath.Length).TrimStart(Path.DirectorySeparatorChar);
                return string.IsNullOrEmpty(relativePath) ? "[Root]" : relativePath;
            }
            
            // Return the full path if it's not under the source path
            return fullPath;
        }

        /// <summary>
        /// Sends an error notification email.
        /// </summary>
        private async Task SendErrorNotificationAsync(string pdfPath, Exception exception)
        {
            StringBuilder bodyBuilder = new StringBuilder();
            bodyBuilder.AppendLine("<h2 style='color: red;'>PDF Conversion Error Notification</h2>");
            bodyBuilder.AppendLine("<div style='font-family: Arial, sans-serif; padding: 15px;'>");
            bodyBuilder.AppendLine("<table style='border-collapse: collapse; width: 100%;'>");
            bodyBuilder.AppendLine("<tr style='background-color: #f2f2f2;'><th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Item</th><th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Detail</th></tr>");
            
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Error Time</td><td style='padding: 10px; border: 1px solid #ddd;'>{DateTime.Now:yyyy-MM-dd HH:mm:ss}</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>PDF File Path</td><td style='padding: 10px; border: 1px solid #ddd;'>{pdfPath}</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Error Type</td><td style='padding: 10px; border: 1px solid #ddd;'>{exception.GetType().Name}</td></tr>");
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Error Message</td><td style='padding: 10px; border: 1px solid #ddd;'>{exception.Message}</td></tr>");
            
            // Add stack trace information
            bodyBuilder.AppendLine($"<tr><td style='padding: 10px; border: 1px solid #ddd;'>Stack Trace</td><td style='padding: 10px; border: 1px solid #ddd; font-family: monospace; white-space: pre-wrap;'>{exception.StackTrace}</td></tr>");
            
            bodyBuilder.AppendLine("</table>");
            bodyBuilder.AppendLine("</div>");

            string subject = $"PDF Conversion Error - {Path.GetFileName(pdfPath)}";
            await _emailService.SendEmailAsync(subject, bodyBuilder.ToString(), true);
            _logger.LogInformation($"Sent error notification email - {Path.GetFileName(pdfPath)}");
        }

        /// <summary>
        /// Ensures that a OneDrive file is fully downloaded before attempting to process it.
        /// OneDrive files may exist as placeholders and need to be downloaded on first access.
        /// </summary>
        private async Task EnsureFileIsDownloadedAsync(string filePath, CancellationToken cancellationToken)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                
                // Check if file exists
                if (!fileInfo.Exists)
                {
                    throw new FileNotFoundException($"File not found: {filePath}");
                }
                
                // Check if this is a OneDrive placeholder file (reparse point)
                bool isOnlinePlaceholder = fileInfo.Attributes.HasFlag(System.IO.FileAttributes.ReparsePoint);
                
                if (isOnlinePlaceholder)
                {
                    _logger.LogInformation($"OneDrive placeholder detected, triggering download: {Path.GetFileName(filePath)}");
                    await TriggerOneDriveDownloadAsync(filePath, cancellationToken);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to ensure file is downloaded: {filePath}");
                throw;
            }
        }

        /// <summary>
        /// Triggers OneDrive download using StorageFile.GetBasicPropertiesAsync() which triggers download.
        /// This is the recommended approach for handling OneDrive Files On-Demand.
        /// </summary>
        private async Task TriggerOneDriveDownloadAsync(string filePath, CancellationToken cancellationToken)
        {
            const int maxRetries = 3;
            const int baseDelayMs = 1000;
            
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    // Method 1: Use StorageFile.GetBasicPropertiesAsync to trigger download
                    try
                    {
                        StorageFile file = await StorageFile.GetFileFromPathAsync(filePath);
                        BasicProperties properties = await file.GetBasicPropertiesAsync(); // 触发下载
                        
                        // Check if file status changed after accessing properties
                        var fileInfo = new FileInfo(filePath);
                        fileInfo.Refresh();
                        
                        if (!fileInfo.Attributes.HasFlag(System.IO.FileAttributes.ReparsePoint))
                        {
                            _logger.LogInformation($"OneDrive file successfully downloaded: {Path.GetFileName(filePath)}");
                        }
                        return; // Success - properties access enables file access
                    }
                    catch (Exception storageEx)
                    {
                        _logger.LogWarning($"StorageFile approach failed, using fallback method: {storageEx.Message}");
                        
                        // Fallback: Traditional file access
                        var fileInfo = new FileInfo(filePath);
                        fileInfo.Refresh();
                        
                        // Access file properties to trigger download
                        var attributes = fileInfo.Attributes;
                        var length = fileInfo.Length;
                        
                        // Try to open file handle to trigger download
                        using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            var streamLength = fileStream.Length;
                            
                            // Try to read first byte to ensure download
                            if (streamLength > 0)
                            {
                                var buffer = new byte[1];
                                _ = await fileStream.ReadAsync(buffer.AsMemory(0, 1), cancellationToken);
                            }
                        }
                    }
                    
                    // Short delay to allow OneDrive to process
                    await Task.Delay(2000, cancellationToken);
                    
                    // Final check
                    var finalFileInfo = new FileInfo(filePath);
                    finalFileInfo.Refresh();
                    var isStillPlaceholder = finalFileInfo.Attributes.HasFlag(System.IO.FileAttributes.ReparsePoint);
                    
                    if (!isStillPlaceholder)
                    {
                        _logger.LogInformation($"OneDrive file successfully downloaded: {Path.GetFileName(filePath)}");
                    }
                    return; // Even if still placeholder, we've triggered access
                }
                catch (Exception ex) when (attempt < maxRetries - 1)
                {
                    _logger.LogWarning($"OneDrive download attempt {attempt + 1} failed, retrying: {ex.Message}");
                    await Task.Delay(baseDelayMs * (attempt + 1), cancellationToken);
                    continue;
                }
                catch (Exception)
                {
                    if (attempt == maxRetries - 1)
                    {
                        // Don't throw - let the actual processing attempt handle the file access
                        _logger.LogWarning($"Could not trigger OneDrive download after {maxRetries} attempts: {Path.GetFileName(filePath)}");
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Validates if a file is a valid PDF format by checking the file header.
        /// </summary>
        private bool IsValidPdfFile(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                
                // Check file size - PDF files should be at least a few bytes
                if (!fileInfo.Exists || fileInfo.Length < 8)
                {
                    return false;
                }
                
                // Check PDF header - should start with "%PDF-"
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var header = new byte[5];
                    int bytesRead = fileStream.Read(header, 0, 5);
                    
                    if (bytesRead < 5)
                    {
                        return false;
                    }
                    
                    // PDF files start with "%PDF-"
                    return header[0] == 0x25 && // %
                           header[1] == 0x50 && // P
                           header[2] == 0x44 && // D
                           header[3] == 0x46 && // F
                           header[4] == 0x2D;   // -
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error validating PDF file format: {filePath}. Error: {ex.Message}");
                return false; // If we can't validate, assume it might be invalid
            }
        }

        /// <summary>
        /// Moves a corrupted or invalid PDF file to the .broken directory to prevent future processing attempts.
        /// </summary>
        private async Task MoveBrokenFileAsync(string pdfPath, string reason)
        {
            try
            {
                var pdfDirectory = Path.GetDirectoryName(pdfPath);
                if (string.IsNullOrEmpty(pdfDirectory))
                {
                    _logger.LogError($"Cannot determine directory for broken file: {pdfPath}");
                    return;
                }

                var brokenDirectory = Path.Combine(pdfDirectory, ".broken");
                Directory.CreateDirectory(brokenDirectory);

                var fileName = Path.GetFileName(pdfPath);
                var brokenFilePath = Path.Combine(brokenDirectory, fileName);

                // If a file with the same name already exists, add a timestamp to make it unique
                int attempt = 1;
                while (File.Exists(brokenFilePath))
                {
                    var fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    var fileExt = Path.GetExtension(fileName);
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var newFileName = $"{fileNameWithoutExt}_{timestamp}_{attempt:000}{fileExt}";
                    brokenFilePath = Path.Combine(brokenDirectory, newFileName);
                    attempt++;
                    
                    // Safety check to prevent infinite loop
                    if (attempt > 999)
                    {
                        _logger.LogError($"Unable to create unique filename for broken file after 999 attempts: {pdfPath}");
                        return;
                    }
                }

                // Move the file to the .broken directory
                File.Move(pdfPath, brokenFilePath);
                _logger.LogWarning($"Moved broken PDF file to .broken directory: {Path.GetFileName(brokenFilePath)}. Reason: {reason}");

                // Create a text file with the error reason
                var reasonFileName = $"{Path.GetFileNameWithoutExtension(brokenFilePath)}_reason.txt";
                var reasonFilePath = Path.Combine(brokenDirectory, reasonFileName);
                var reasonContent = $"Moved on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\nOriginal file: {Path.GetFileName(pdfPath)}\nReason: {reason}";
                await File.WriteAllTextAsync(reasonFilePath, reasonContent);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to move broken file to .broken directory: {pdfPath}");
                
                // If moving fails, at least try to log the issue to a central broken files log
                try
                {
                    var logDirectory = Path.GetDirectoryName(pdfPath);
                    if (!string.IsNullOrEmpty(logDirectory))
                    {
                        var logFilePath = Path.Combine(logDirectory, "broken_files_log.txt");
                        var logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Failed to move: {pdfPath} - Reason: {reason} - Error: {ex.Message}\n";
                        await File.AppendAllTextAsync(logFilePath, logEntry);
                    }
                }
                catch
                {
                    // If even logging fails, we've done our best
                    _logger.LogError($"Could not log broken file information to broken_files_log.txt for: {pdfPath}");
                }
            }
        }

        /// <summary>
        /// Moves a successfully processed PDF file to the .pdf archive directory, avoiding overwrites by renaming if necessary.
        /// </summary>
        private void MoveToArchive(string pdfPath, string pdfDirectory)
        {
            try
            {
                var archiveDirectory = Path.Combine(pdfDirectory, ".pdf");
                Directory.CreateDirectory(archiveDirectory);

                var fileName = Path.GetFileName(pdfPath);
                var archiveFilePath = Path.Combine(archiveDirectory, fileName);

                // If a file with the same name already exists, add a timestamp to make it unique
                int attempt = 1;
                while (File.Exists(archiveFilePath))
                {
                    var fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    var fileExt = Path.GetExtension(fileName);
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var newFileName = $"{fileNameWithoutExt}_{timestamp}_{attempt:000}{fileExt}";
                    archiveFilePath = Path.Combine(archiveDirectory, newFileName);
                    attempt++;
                    
                    // Safety check to prevent infinite loop
                    if (attempt > 999)
                    {
                        _logger.LogError($"Unable to create unique filename for archive after 999 attempts: {pdfPath}");
                        // Fallback: use the current timestamp and milliseconds
                        var fallbackName = $"{fileNameWithoutExt}_{DateTime.Now:yyyyMMdd_HHmmss_fff}{fileExt}";
                        archiveFilePath = Path.Combine(archiveDirectory, fallbackName);
                        break;
                    }
                }

                // Move the file to the archive directory
                File.Move(pdfPath, archiveFilePath);
                _logger.LogInformation($"Archived processed PDF file: {Path.GetFileName(archiveFilePath)}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to move processed file to .pdf archive directory: {pdfPath}");
                
                // Re-throw the exception as this is a critical failure
                throw;
            }
        }

    }
}
