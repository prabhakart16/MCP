using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace ExcelMcpServer.Client;

/// <summary>
/// Interactive client that allows users to ask natural language questions
/// </summary>
public class InteractiveClient
{
    private readonly Process _serverProcess;
    private readonly StreamWriter _stdin;
    private readonly StreamReader _stdout;
    private readonly JsonSerializerOptions _jsonOptions;
    private int _requestId = 0;
    private bool _isInitialized = false;

    public InteractiveClient(string serverExecutablePath, string excelFilePath)
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
                Arguments = excelFilePath,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };

        // Capture stderr for server logs (optional - can suppress)
        _serverProcess.ErrorDataReceived += (sender, e) =>
        {
            if (!string.IsNullOrEmpty(e.Data) && e.Data.Contains("ERROR"))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"[Server Error] {e.Data}");
                Console.ResetColor();
            }
        };

        _serverProcess.Start();
        _serverProcess.BeginErrorReadLine();

        _stdin = _serverProcess.StandardInput;
        _stdout = _serverProcess.StandardOutput;
    }

    public async Task InitializeAsync()
    {
        var request = new
        {
            jsonrpc = "2.0",
            id = (++_requestId).ToString(),
            method = "initialize",
            @params = new
            {
                protocolVersion = "2024-11-05",
                clientInfo = new { name = "InteractiveClient", version = "1.0.0" },
                capabilities = new { }
            }
        };

        var response = await SendRequestAsync(request);
        _isInitialized = true;

        // Wait a moment for server to fully initialize
        await Task.Delay(500);
    }

    public async Task<QueryResult> AskQuestionAsync(string naturalLanguageQuery, int limit = 20)
    {
        if (!_isInitialized)
            throw new InvalidOperationException("Client not initialized. Call InitializeAsync first.");

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
                    query = naturalLanguageQuery,
                    limit,
                    skip = 0
                }
            }
        };

        var response = await SendRequestAsync(request);
        return ParseQueryResult(response);
    }

    public async Task<Statistics> GetStatisticsAsync()
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

        var response = await SendRequestAsync(request);
        return ParseStatistics(response);
    }

    private async Task<JsonDocument> SendRequestAsync(object request)
    {
        var json = JsonSerializer.Serialize(request, _jsonOptions);
        await _stdin.WriteLineAsync(json);
        await _stdin.FlushAsync();

        var response = await _stdout.ReadLineAsync();
        if (string.IsNullOrEmpty(response))
            throw new Exception("Empty response from server");

        return JsonDocument.Parse(response);
    }

    private QueryResult ParseQueryResult(JsonDocument doc)
    {
        var result = new QueryResult();

        if (doc.RootElement.TryGetProperty("error", out var error))
        {
            result.Error = error.GetProperty("message").GetString() ?? "Unknown error";
            return result;
        }

        var content = doc.RootElement
            .GetProperty("result")
            .GetProperty("content")[0]
            .GetProperty("text")
            .GetString();

        if (string.IsNullOrEmpty(content))
            return result;

        var queryResponse = JsonSerializer.Deserialize<QueryResponse>(content, _jsonOptions);
        if (queryResponse == null)
            return result;

        result.Success = queryResponse.Success;
        result.Message = queryResponse.Message;
        result.TotalCount = queryResponse.TotalCount;
        result.Loans = queryResponse.Data;
        result.QueryType = queryResponse.Metadata?.QueryType;
        result.ExecutionTimeMs = queryResponse.Metadata?.ExecutionTimeMs ?? 0;
        result.Stats = queryResponse.Metadata?.Statistics;

        return result;
    }

    private Statistics ParseStatistics(JsonDocument doc)
    {
        var content = doc.RootElement
            .GetProperty("result")
            .GetProperty("content")[0]
            .GetProperty("text")
            .GetString();

        return JsonSerializer.Deserialize<Statistics>(content!, _jsonOptions)
            ?? new Statistics();
    }

    public void Dispose()
    {
        _stdin?.Dispose();
        _stdout?.Dispose();

        if (!_serverProcess.HasExited)
            _serverProcess.Kill();

        _serverProcess?.Dispose();
    }
}

// Result models
public class QueryResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public string? Error { get; set; }
    public int TotalCount { get; set; }
    public List<LoanRecord> Loans { get; set; } = new();
    public string? QueryType { get; set; }
    public long ExecutionTimeMs { get; set; }
    public Dictionary<string, object>? Stats { get; set; }
}

public class QueryResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<LoanRecord> Data { get; set; } = new();
    public int TotalCount { get; set; }
    public QueryMetadata? Metadata { get; set; }
}

public class QueryMetadata
{
    public string QueryType { get; set; } = string.Empty;
    public long ExecutionTimeMs { get; set; }
    public Dictionary<string, object> Statistics { get; set; } = new();
}

public class LoanRecord
{
    public string LoanID { get; set; } = string.Empty;
    public string BorrowerName { get; set; } = string.Empty;
    public decimal Servicer_LoanAmount { get; set; }
    public decimal FNMA_LoanAmount { get; set; }
    public decimal DifferenceAmount { get; set; }
    public string ReconciledStatus { get; set; } = string.Empty;
}

public class Statistics
{
    public int TotalRecords { get; set; }
    public DateTime LastLoadTime { get; set; }
    public bool DataLoaded { get; set; }
}

// Interactive Console Program
public class InteractiveProgram
{
    public static async Task Main(string[] args)
    {
        Console.Clear();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("╔════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║     Excel MCP Server - Natural Language Query Interface   ║");
        Console.WriteLine("╚════════════════════════════════════════════════════════════╝");
        Console.ResetColor();
        Console.WriteLine();

        if (args.Length < 2)
        {
            Console.WriteLine("Usage: InteractiveClient <server-exe> <excel-file>");
            Console.WriteLine("Example: InteractiveClient ExcelMcpServer.exe loans.xlsx");
            return;
        }

        var serverPath = args[0];
        var excelPath = args[1];

        Console.WriteLine($"📁 Excel File: {Path.GetFileName(excelPath)}");
        Console.WriteLine($"🖥️  Server: {Path.GetFileName(serverPath)}");
        Console.WriteLine();
        Console.Write("⏳ Initializing connection...");

        InteractiveClient? client = null;
        try
        {
            client = new InteractiveClient(serverPath, excelPath);
            await client.InitializeAsync();

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(" ✓ Ready!");
            Console.ResetColor();
            Console.WriteLine();

            // Show statistics
            var stats = await client.GetStatisticsAsync();
            Console.WriteLine($"📊 Loaded {stats.TotalRecords:N0} loan records");
            Console.WriteLine($"🕐 Last updated: {stats.LastLoadTime:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine();

            ShowHelp();

            // Main interaction loop
            while (true)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write("❓ Your question (or 'help', 'examples', 'stats', 'exit'): ");
                Console.ResetColor();

                var input = Console.ReadLine()?.Trim();
                if (string.IsNullOrEmpty(input))
                    continue;

                if (input.Equals("exit", StringComparison.OrdinalIgnoreCase) ||
                    input.Equals("quit", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("👋 Goodbye!");
                    break;
                }

                if (input.Equals("help", StringComparison.OrdinalIgnoreCase))
                {
                    ShowHelp();
                    continue;
                }

                if (input.Equals("examples", StringComparison.OrdinalIgnoreCase))
                {
                    ShowExamples();
                    continue;
                }

                if (input.Equals("stats", StringComparison.OrdinalIgnoreCase))
                {
                    var s = await client.GetStatisticsAsync();
                    Console.WriteLine($"\n📊 Dataset Statistics:");
                    Console.WriteLine($"   Total Records: {s.TotalRecords:N0}");
                    Console.WriteLine($"   Last Updated: {s.LastLoadTime:yyyy-MM-dd HH:mm:ss}");
                    continue;
                }

                // Process natural language query
                await ProcessQueryAsync(client, input);
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n❌ Error: {ex.Message}");
            Console.ResetColor();
        }
        finally
        {
            client?.Dispose();
        }
    }

    private static async Task ProcessQueryAsync(InteractiveClient client, string query)
    {
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine($"⏳ Processing query...");
        Console.ResetColor();

        try
        {
            var result = await client.AskQuestionAsync(query, limit: 20);

            if (!string.IsNullOrEmpty(result.Error))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"❌ Error: {result.Error}");
                Console.ResetColor();
                return;
            }

            if (!result.Success)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"⚠️  {result.Message}");
                Console.ResetColor();
                return;
            }

            // Display results
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"✓ {result.Message}");
            Console.ResetColor();
            Console.WriteLine($"   Query Type: {result.QueryType}");
            Console.WriteLine($"   Execution Time: {result.ExecutionTimeMs}ms");
            Console.WriteLine();

            if (result.Loans.Any())
            {
                Console.WriteLine($"📋 Showing {result.Loans.Count} of {result.TotalCount:N0} results:");
                Console.WriteLine();

                var table = new ConsoleTable();
                table.AddHeaders("Loan ID", "Borrower", "Servicer Amt", "FNMA Amt", "Difference", "Status");

                foreach (var loan in result.Loans)
                {
                    table.AddRow(
                        loan.LoanID,
                        loan.BorrowerName.Length > 20
                            ? loan.BorrowerName.Substring(0, 17) + "..."
                            : loan.BorrowerName,
                        FormatCurrency(loan.Servicer_LoanAmount),
                        FormatCurrency(loan.FNMA_LoanAmount),
                        FormatCurrency(loan.DifferenceAmount, true),
                        loan.ReconciledStatus
                    );
                }

                table.Display();
            }

            // Show statistics if available
            if (result.Stats != null && result.Stats.Any())
            {
                Console.WriteLine();
                Console.WriteLine("📊 Statistics:");
                foreach (var stat in result.Stats)
                {
                    if (stat.Value is decimal dec)
                        Console.WriteLine($"   {stat.Key}: {FormatCurrency(dec)}");
                    else
                        Console.WriteLine($"   {stat.Key}: {stat.Value}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"❌ Error processing query: {ex.Message}");
            Console.ResetColor();
        }
    }

    private static string FormatCurrency(decimal value, bool colorize = false)
    {
        var formatted = value.ToString("C2");
        if (!colorize || value == 0)
            return formatted;

        return value > 0
            ? $"+{formatted}"
            : formatted;
    }

    private static void ShowHelp()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("═══════════════════ HELP ═══════════════════");
        Console.ResetColor();
        Console.WriteLine();
        Console.WriteLine("You can ask questions in natural language. Examples:");
        Console.WriteLine();
        Console.WriteLine("  • Find all mismatches");
        Console.WriteLine("  • Show loans where difference > 5000");
        Console.WriteLine("  • List unreconciled loans");
        Console.WriteLine("  • Find loan ABC-123");
        Console.WriteLine("  • Search borrower John Smith");
        Console.WriteLine();
        Console.WriteLine("Commands:");
        Console.WriteLine("  help      - Show this help message");
        Console.WriteLine("  examples  - Show more query examples");
        Console.WriteLine("  stats     - Show dataset statistics");
        Console.WriteLine("  exit      - Exit the program");
        Console.WriteLine();
    }

    private static void ShowExamples()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("═══════════════ QUERY EXAMPLES ═══════════════");
        Console.ResetColor();
        Console.WriteLine();

        var examples = new[]
        {
            ("Mismatch Queries", new[]
            {
                "Find all mismatches",
                "Show differences between servicer and FNMA",
                "Reconcile amounts and find issues"
            }),
            ("Threshold Queries", new[]
            {
                "Show loans where difference > 10000",
                "Find loans with difference < -5000",
                "List loans where DifferenceAmount > 2500"
            }),
            ("Status Queries", new[]
            {
                "Show reconciled loans",
                "List unreconciled loans",
                "Find pending reconciliation"
            }),
            ("Search Queries", new[]
            {
                "Find loan LN-001234",
                "Search borrower John",
                "Look up borrower name Smith"
            }),
            ("General Queries", new[]
            {
                "List all loans",
                "Show everything",
                "Give me all records"
            })
        };

        foreach (var (category, queries) in examples)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"📁 {category}:");
            Console.ResetColor();
            foreach (var query in queries)
            {
                Console.WriteLine($"   • {query}");
            }
            Console.WriteLine();
        }
    }
}

// Simple console table formatter
public class ConsoleTable
{
    private readonly List<string> _headers = new();
    private readonly List<List<string>> _rows = new();

    public void AddHeaders(params string[] headers)
    {
        _headers.AddRange(headers);
    }

    public void AddRow(params object[] values)
    {
        _rows.Add(values.Select(v => v?.ToString() ?? "").ToList());
    }

    public void Display()
    {
        if (!_headers.Any()) return;

        var columnWidths = new int[_headers.Count];

        // Calculate column widths
        for (int i = 0; i < _headers.Count; i++)
        {
            columnWidths[i] = _headers[i].Length;
            foreach (var row in _rows)
            {
                if (i < row.Count && row[i].Length > columnWidths[i])
                    columnWidths[i] = row[i].Length;
            }
            columnWidths[i] += 2; // Padding
        }

        // Print header
        Console.ForegroundColor = ConsoleColor.Cyan;
        PrintRow(_headers, columnWidths);
        Console.WriteLine(new string('─', columnWidths.Sum() + _headers.Count + 1));
        Console.ResetColor();

        // Print rows
        foreach (var row in _rows)
        {
            PrintRow(row, columnWidths);
        }
    }

    private void PrintRow(List<string> values, int[] widths)
    {
        Console.Write("│");
        for (int i = 0; i < values.Count; i++)
        {
            Console.Write($" {values[i].PadRight(widths[i] - 1)}│");
        }
        Console.WriteLine();
    }
}