using PDFtoImage;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                    .Where(p => !Path.GetDirectoryName(p).EndsWith(Path.DirectorySeparatorChar + ".pdf") && Path.GetDirectoryName(p)!= "Z1-template")
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
                        
                        // Create a .pdf archive folder in the PDF file's directory
                        var archiveDirectory = Path.Combine(pdfDirectory, ".pdf");
                        Directory.CreateDirectory(archiveDirectory);

                        //Create a .archive folder in the PDF file's directory
                        var archivePath = Path.Combine(pdfDirectory, ".archive");
                        if(!Directory.Exists(archivePath))
                        {
                            Directory.CreateDirectory(archivePath);
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
            // Get the number of pages with limited memory usage
            int pageCount;
            using (var pdfStream = File.OpenRead(pdfPath))
            {
                pageCount = Conversion.GetPageCount(pdfStream);
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

                // Open streams for only the current page
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
            string destPdf = Path.Combine(pdfDirectory, ".pdf", Path.GetFileName(pdfPath));
            File.Move(pdfPath, destPdf);
            
            // Update directory stats for successful conversion
            var currentStats = directoryStats[pdfDirectory];
            directoryStats[pdfDirectory] = (currentStats.TotalFiles, 
                                           currentStats.SuccessFiles + 1, 
                                           currentStats.FailedFiles, 
                                           currentStats.TotalPages + pageCount);
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
    }
}
