using ExcelMcpServer.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace ExcelMcpServer;

class Program
{
    static async Task<int> Main(string[] args)
    {
        try
        {
            var builder = Host.CreateApplicationBuilder(args);

            // Configure logging - CRITICAL: MCP servers must log to STDERR only
            builder.Logging.ClearProviders();
            builder.Logging.AddSimpleConsole(options =>
            {
                options.SingleLine = true;
            });
            builder.Logging.AddConsole(options =>
            {
                // Force all logs to stderr for MCP compatibility
                options.LogToStandardErrorThreshold = LogLevel.Trace;
            });
            builder.Logging.SetMinimumLevel(LogLevel.Information);

            // Register services
            builder.Services.AddSingleton<ExcelDataService>();
            builder.Services.AddSingleton<McpServerService>();
            builder.Services.AddHostedService<McpServerHostedService>();

            var host = builder.Build();

            // Load Excel file on startup
            var excelService = host.Services.GetRequiredService<ExcelDataService>();
            var excelFilePath = @"C:\Prabhakar\MCP\sampleFile\LoanReconciliationSample.xlsx";

            var loaded = await excelService.LoadExcelFileAsync(excelFilePath);
            if (!loaded)
            {
                Console.Error.WriteLine("Failed to load Excel file");
                return 1;
            }

            await host.RunAsync();
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Fatal error: {ex.Message}");
            return 1;
        }
    }
}

// Hosted service to run MCP server
public class McpServerHostedService : IHostedService
{
    private readonly McpServerService _mcpServer;
    private readonly ILogger<McpServerHostedService> _logger;
    private CancellationTokenSource? _cts;

    public McpServerHostedService(
        McpServerService mcpServer,
        ILogger<McpServerHostedService> logger)
    {
        _mcpServer = mcpServer;
        _logger = logger;
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("MCP Server Hosted Service starting");
        _cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        _ = Task.Run(async () =>
        {
            try
            {
                await _mcpServer.StartAsync(_cts.Token);
            }
            catch (Exception ex)
            {
                _logger.LogCritical(ex, "MCP Server crashed");
            }
        }, _cts.Token);

        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("MCP Server Hosted Service stopping");
        _cts?.Cancel();
        return Task.CompletedTask;
    }
}