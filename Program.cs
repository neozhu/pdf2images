using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Threading.Tasks;
using pdf2images;
using Serilog;
using Microsoft.Extensions.Configuration;

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

    var builder = Host.CreateApplicationBuilder(args);
    
    // Use Serilog as the logging provider
    builder.Services.AddSerilog();
    
    // Configure Windows Service support
    builder.Services.AddWindowsService(options =>
    {
        options.ServiceName = "PDF2Images Service";
    });
    
    // Register services for dependency injection
    builder.Services.AddSingleton<EmailService>(provider => 
        new EmailService(builder.Configuration));
        
    builder.Services.AddSingleton<Pdf2ImageService>(provider =>
        new Pdf2ImageService(
            builder.Configuration,
            provider.GetRequiredService<EmailService>(),
            provider.GetRequiredService<ILogger<Pdf2ImageService>>()));
            
    builder.Services.AddHostedService<Worker>();

    // Build the service provider
    var host = builder.Build();
    
    // Get the services
    var serviceProvider = host.Services;
    var emailService = serviceProvider.GetRequiredService<EmailService>();
    var logger = serviceProvider.GetRequiredService<ILogger<Pdf2ImageService>>();
    var pdf2ImageService = serviceProvider.GetRequiredService<Pdf2ImageService>();
    // Run the host
    await host.RunAsync();
}
catch (Exception ex)
{
    // Handle startup exceptions
    Log.Fatal(ex, "Application startup failed");
    await SendExceptionEmailAsync(ex, "Application Startup Error");
    
    // Re-throw to terminate the application if it can't start properly
    throw;
}
finally
{
    Log.CloseAndFlush();
}


