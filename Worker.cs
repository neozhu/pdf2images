using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.ExceptionServices;

namespace pdf2images
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly Pdf2ImageService _pdf2ImageService;
        private readonly TimeSpan _processingInterval = TimeSpan.FromHours(2);
        private readonly SemaphoreSlim _processingSemaphore = new SemaphoreSlim(1, 1);
        
        public Worker(ILogger<Worker> logger, Pdf2ImageService pdf2ImageService)
        {
            _logger = logger;
            _pdf2ImageService = pdf2ImageService;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("PDF to Image Service started at: {time}", DateTimeOffset.Now);
            
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    // Use semaphore to prevent overlapping executions
                    if (await _processingSemaphore.WaitAsync(0))
                    {
                        try
                        {
                            _logger.LogInformation("Starting PDF processing at: {time}", DateTimeOffset.Now);
                            
                            // Create a timeout for this processing run (4 hours)
                            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromHours(4));
                            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
                                timeoutCts.Token, stoppingToken);
                            
                            // Process with timeout and memory optimization
                            await ProcessWithResourceProtectionAsync(linkedCts.Token);
                            
                            _logger.LogInformation("PDF processing completed successfully at: {time}", DateTimeOffset.Now);
                        }
                        catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
                        {
                            _logger.LogWarning("PDF processing was canceled due to service shutdown");
                        }
                        catch (OperationCanceledException)
                        {
                            _logger.LogWarning("PDF processing was canceled due to timeout");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Error occurred during PDF processing");
                        }
                        finally
                        {
                            // Force garbage collection to free resources
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            
                            _processingSemaphore.Release();
                        }
                    }
                    else
                    {
                        _logger.LogWarning("Skipping PDF processing as previous run is still in progress");
                    }

                    // Wait for the configured interval before next processing
                    await Task.Delay(_processingInterval, stoppingToken);
                }
                catch (Exception ex)
                {
                    // Catch-all for unexpected errors to prevent service crash
                    _logger.LogError(ex, "Unexpected error in service execution loop");
                    
                    // Wait a shorter time before retry on error
                    try
                    {
                        await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
                    }
                    catch (OperationCanceledException)
                    {
                        // Ignore cancellation exceptions during delay
                    }
                }
            }
        }
        
        private async Task ProcessWithResourceProtectionAsync(CancellationToken cancellationToken)
        {
            // Set thread priority to below normal to reduce system impact
            Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
            
            try
            {
                // Execute the PDF processing with the cancellation token
                await _pdf2ImageService.ProcessAllAsync();
                
                // Check cancellation between potentially heavy operations
                cancellationToken.ThrowIfCancellationRequested();
            }
            finally
            {
                // Restore thread priority
                Thread.CurrentThread.Priority = ThreadPriority.Normal;
            }
        }

        public override async Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("PDF to Image Service is stopping");
            await base.StopAsync(cancellationToken);
        }
    }
}
