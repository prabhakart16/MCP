using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace ExcelMcpServer.Client;

/// <summary>
/// Example MCP client for testing the Excel MCP Server
/// </summary>
public class McpClient : IDisposable
{
    private readonly Process _serverProcess;
    private readonly StreamWriter _stdin;
    private readonly StreamReader _stdout;
    private readonly JsonSerializerOptions _jsonOptions;
    private int _requestId = 0;

    public McpClient(string serverExecutablePath)
    {
        _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        _serverProcess = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = serverExecutablePath,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };

        _serverProcess.Start();
        _stdin = _serverProcess.StandardInput;
        _stdout = _serverProcess.StandardOutput;
    }

    public async Task<JsonDocument> InitializeAsync()
    {
        var request = new
        {
            jsonrpc = "2.0",
            id = (++_requestId).ToString(),
            method = "initialize",
            @params = new
            {
                protocolVersion = "2024-11-05",
                clientInfo = new
                {
                    name = "ExcelMcpClient",
                    version = "1.0.0"
                },
                capabilities = new { }
            }
        };

        return await SendRequestAsync(request);
    }

    public async Task<JsonDocument> QueryLoansAsync(
        string query,
        int? limit = null,
        int? skip = null)
    {
        var request = new
        {
            jsonrpc = "2.0",
            id = (++_requestId).ToString(),
            method = "tools/call",
            @params = new
            {
                name = "query_loans",
                arguments = new
                {
                    query,
                    limit,
                    skip
                }
            }
        };

        return await SendRequestAsync(request);
    }

    public async Task<JsonDocument> GetStatisticsAsync()
    {
        var request = new
        {
            jsonrpc = "2.0",
            id = (++_requestId).ToString(),
            method = "tools/call",
            @params = new
            {
                name = "get_statistics",
                arguments = new { }
            }
        };

        return await SendRequestAsync(request);
    }

    private async Task<JsonDocument> SendRequestAsync(object request)
    {
        var json = JsonSerializer.Serialize(request, _jsonOptions);
        await _stdin.WriteLineAsync(json);
        await _stdin.FlushAsync();

        string? response;
        // Skip lines that are not valid JSON (e.g., server logs or errors)
        while (true)
        {
            response = await _stdout.ReadLineAsync();
            if (string.IsNullOrEmpty(response))
                throw new Exception("Empty response from server");

            // Try to detect if the line is likely JSON (starts with '{' or '[')
            var trimmed = response.TrimStart();
            if (trimmed.StartsWith("{") || trimmed.StartsWith("["))
                break;
            // Optionally, log or handle non-JSON lines here
        }

        return JsonDocument.Parse(response);
    }

    public void Dispose()
    {
        _stdin?.Dispose();
        _stdout?.Dispose();

        if (!_serverProcess.HasExited)
        {
            _serverProcess.Kill();
        }

        _serverProcess?.Dispose();
    }
}

// Example usage program
public class ClientExample
{
    public static async Task Main(string[] args)
    {
        var serverPath = args.Length > 0 ? args[0] : "./Recon.exe";

        using var client = new McpClient(serverPath);

        Console.WriteLine("=== Initializing MCP Client ===");
        var initResponse = await client.InitializeAsync();
        Console.WriteLine(initResponse.RootElement.GetRawText());
        Console.WriteLine();

        // Example 1: Find all mismatches
        Console.WriteLine("=== Query 1: Find Mismatches ===");
        var query1 = await client.QueryLoansAsync(
            "Reconcile Servicer_LoanAmount and FNMA_LoanAmount and find mismatches",
            limit: 10);
        Console.WriteLine(FormatResponse(query1));
        Console.WriteLine();

        // Example 2: Find loans with difference > 5000
        Console.WriteLine("=== Query 2: Difference > 5000 ===");
        var query2 = await client.QueryLoansAsync(
            "Show loans where DifferenceAmount > 5000",
            limit: 5);
        Console.WriteLine(FormatResponse(query2));
        Console.WriteLine();

        // Example 3: Find unreconciled loans
        Console.WriteLine("=== Query 3: Unreconciled Loans ===");
        var query3 = await client.QueryLoansAsync(
            "Show unreconciled loans",
            limit: 10);
        Console.WriteLine(FormatResponse(query3));
        Console.WriteLine();

        // Example 4: Get statistics
        Console.WriteLine("=== Query 4: Dataset Statistics ===");
        var stats = await client.GetStatisticsAsync();
        Console.WriteLine(stats.RootElement.GetRawText());
    }

    private static string FormatResponse(JsonDocument doc)
    {
        var options = new JsonSerializerOptions { WriteIndented = true };
        return JsonSerializer.Serialize(doc.RootElement, options);
    }
}