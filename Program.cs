using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Threading;
using System.Threading.Tasks;
using pdf2images;
using Serilog;

// Configure unhandled exception handler for the entire application
AppDomain.CurrentDomain.UnhandledException += (sender, args) => 
{
    var exception = args.ExceptionObject as Exception;
    if (exception != null)
        HandleGlobalException(exception, "Unhandled Exception");
};

// Configure tasks to observe unhandled exceptions
TaskScheduler.UnobservedTaskException += (sender, args) =>
{
    HandleGlobalException(args.Exception, "Unobserved Task Exception");
    args.SetObserved(); // Prevent process termination
};

static void HandleGlobalException(Exception ex, string type)
{
    try
    {
        // Create a simple logger to log the global exception
        var loggerFactory = LoggerFactory.Create(logging => logging.AddConsole());
        var logger = loggerFactory.CreateLogger("GlobalExceptionHandler");
        logger.LogCritical(ex, $"{type} occurred: {ex.Message}");
        
        // Try to send email about the exception
        SendExceptionEmailAsync(ex, type).Wait();
    }
    catch
    {
        // Last resort - we can't do much if this fails
    }
}

static async Task SendExceptionEmailAsync(Exception ex, string type)
{
    try
    {
        // Build minimal configuration to create email service
        var config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true)
            .AddEnvironmentVariables()
            .Build();
        
        var emailService = new EmailService(config);
        
        string subject = $"PDF2Images Service Critical Error: {type}";
        string body = $@"
<h2 style='color: #cc0000;'>Critical Error in PDF2Images Service</h2>
<div style='font-family: Arial, sans-serif; padding: 15px;'>
    <table style='border-collapse: collapse; width: 100%;'>
        <tr style='background-color: #f2f2f2;'>
            <th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Error Details</th>
            <th style='padding: 10px; text-align: left; border: 1px solid #ddd;'>Value</th>
        </tr>
        <tr>
            <td style='padding: 10px; border: 1px solid #ddd;'>Error Time</td>
            <td style='padding: 10px; border: 1px solid #ddd;'>{DateTime.Now:yyyy-MM-dd HH:mm:ss}</td>
        </tr>
        <tr>
            <td style='padding: 10px; border: 1px solid #ddd;'>Error Type</td>
            <td style='padding: 10px; border: 1px solid #ddd;'>{type}</td>
        </tr>
        <tr>
            <td style='padding: 10px; border: 1px solid #ddd;'>Exception Type</td>
            <td style='padding: 10px; border: 1px solid #ddd;'>{ex.GetType().FullName}</td>
        </tr>
        <tr>
            <td style='padding: 10px; border: 1px solid #ddd;'>Message</td>
            <td style='padding: 10px; border: 1px solid #ddd;'>{ex.Message}</td>
        </tr>
        <tr>
            <td style='padding: 10px; border: 1px solid #ddd;'>Stack Trace</td>
            <td style='padding: 10px; border: 1px solid #ddd; font-family: monospace; white-space: pre-wrap;'>{ex.StackTrace}</td>
        </tr>
    </table>
    
    <h3 style='margin-top: 20px;'>Inner Exception Details</h3>
    {GetInnerExceptionHtml(ex.InnerException)}
</div>";

        await emailService.SendEmailAsync(subject, body, true);
    }
    catch
    {
        // If this fails, we've already tried our best
    }
}

static string GetInnerExceptionHtml(Exception? innerEx)
{
    if (innerEx == null)
    {
        return "<p>No inner exception</p>";
    }
    
    return $@"
<table style='border-collapse: collapse; width: 100%; margin-top: 10px;'>
    <tr>
        <td style='padding: 10px; border: 1px solid #ddd;'>Inner Exception Type</td>
        <td style='padding: 10px; border: 1px solid #ddd;'>{innerEx.GetType().FullName}</td>
    </tr>
    <tr>
        <td style='padding: 10px; border: 1px solid #ddd;'>Message</td>
        <td style='padding: 10px; border: 1px solid #ddd;'>{innerEx.Message}</td>
    </tr>
    <tr>
        <td style='padding: 10px; border: 1px solid #ddd;'>Stack Trace</td>
        <td style='padding: 10px; border: 1px solid #ddd; font-family: monospace; white-space: pre-wrap;'>{innerEx.StackTrace}</td>
    </tr>
</table>
{GetInnerExceptionHtml(innerEx.InnerException)}";
}

// Main program
try
{
    // Configure Serilog
    Log.Logger = new LoggerConfiguration()
        .MinimumLevel.Information()
        .WriteTo.Console()
        .WriteTo.File(
            path: "logs/pdf2images-.txt",
            rollingInterval: RollingInterval.Day,
            rollOnFileSizeLimit: true,
            fileSizeLimitBytes: 10485760,
            retainedFileCountLimit: 10,
            buffered: false,
            flushToDiskInterval: TimeSpan.FromSeconds(1))
        .CreateLogger();

    // Create configuration
    var configuration = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: false)
        .AddJsonFile("appsettings.Development.json", optional: true)
        .AddEnvironmentVariables()
        .Build();

    // Create logger factory for dependency injection
    using var loggerFactory = LoggerFactory.Create(logging => 
    {
        logging.AddSerilog();
    });

    var logger = loggerFactory.CreateLogger<Pdf2ImageService>();
    
    // Create services directly
    var emailService = new EmailService(configuration);
    var pdf2ImageService = new Pdf2ImageService(configuration, emailService, logger);
    
    Log.Information("Starting PDF to Image conversion process...");
    
    // Create cancellation token for graceful shutdown
    using var cts = new CancellationTokenSource();
    
    // Handle Ctrl+C for graceful shutdown
    Console.CancelKeyPress += (sender, e) =>
    {
        e.Cancel = true; // Prevent immediate termination
        Log.Information("Cancellation requested. Shutting down gracefully...");
        cts.Cancel();
    };
    
    // Run the PDF processing
    await pdf2ImageService.ProcessAllAsync(batchSize: 10, cancellationToken: cts.Token);
    
    Log.Information("PDF to Image conversion completed successfully.");
}
catch (OperationCanceledException)
{
    Log.Information("PDF processing was canceled by user.");
}
catch (Exception ex)
{
    // Handle startup and processing exceptions
    Log.Fatal(ex, "Application failed with error: {ErrorMessage}", ex.Message);
    await SendExceptionEmailAsync(ex, "Application Error");
    
    // Set exit code to indicate error
    Environment.Exit(1);
}
finally
{
    Log.Information("Application shutting down...");
    Log.CloseAndFlush();
}


