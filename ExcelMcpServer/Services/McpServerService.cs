using ExcelMcpServer.Models;
using Microsoft.Extensions.Logging;
using System.Text;
using System.Text.Json;

namespace ExcelMcpServer.Services;

public class McpServerService
{
    private readonly ExcelDataService _dataService;
    private readonly ILogger<McpServerService> _logger;
    private readonly JsonSerializerOptions _jsonOptions;

    public McpServerService(ExcelDataService dataService, ILogger<McpServerService> logger)
    {
        _dataService = dataService;
        _logger = logger;
        _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };
    }

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("MCP Server starting...");

        // Initialize by reading from stdin/stdout (MCP protocol)
        await ProcessMcpMessagesAsync(cancellationToken);
    }

    private async Task ProcessMcpMessagesAsync(CancellationToken cancellationToken)
    {
        try
        {
            using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
            using var writer = new StreamWriter(Console.OpenStandardOutput(), Encoding.UTF8)
            {
                AutoFlush = true
            };

            while (!cancellationToken.IsCancellationRequested)
            {
                var line = await reader.ReadLineAsync(cancellationToken);
                if (string.IsNullOrEmpty(line)) continue;

                try
                {
                    var request = JsonSerializer.Deserialize<McpRequest>(line, _jsonOptions);
                    if (request == null) continue;

                    var response = await HandleMcpRequestAsync(request);
                    var responseJson = JsonSerializer.Serialize(response, _jsonOptions);

                    await writer.WriteLineAsync(responseJson);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing MCP request");
                    var errorResponse = new McpResponse
                    {
                        Id = Guid.NewGuid().ToString(),
                        Error = new McpError
                        {
                            Code = -32603,
                            Message = ex.Message
                        }
                    };
                    await writer.WriteLineAsync(
                        JsonSerializer.Serialize(errorResponse, _jsonOptions));
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogCritical(ex, "Fatal error in MCP message processing");
        }
    }

    private async Task<McpResponse> HandleMcpRequestAsync(McpRequest request)
    {
        var response = new McpResponse { Id = request.Id };

        try
        {
            switch (request.Method)
            {
                case "initialize":
                    response.Result = new
                    {
                        protocolVersion = "2024-11-05",
                        serverInfo = new
                        {
                            name = "ExcelMcpServer",
                            version = "1.0.0"
                        },
                        capabilities = new
                        {
                            tools = new { },
                            resources = new { }
                        }
                    };
                    break;

                case "tools/list":
                    response.Result = new
                    {
                        tools = new[]
                        {
                            new
                            {
                                name = "query_loans",
                                description = "Query loan data with natural language",
                                inputSchema = new
                                {
                                    type = "object",
                                    properties = new
                                    {
                                        query = new { type = "string", description = "Natural language query" },
                                        limit = new { type = "integer", description = "Max results to return" },
                                        skip = new { type = "integer", description = "Records to skip" }
                                    },
                                    required = new[] { "query" }
                                }
                            },
                            new
                            {
                                name = "get_statistics",
                                description = "Get overall dataset statistics",
                                inputSchema = new { type = "object", properties = new { } }
                            }
                        }
                    };
                    break;

                case "tools/call":
                    var toolName = request.Params?.GetProperty("name").GetString();
                    var arguments = request.Params?.GetProperty("arguments");

                    if (toolName == "query_loans")
                    {
                        var queryRequest = JsonSerializer.Deserialize<QueryRequest>(
                            arguments?.GetRawText() ?? "{}", _jsonOptions);

                        var queryResponse = await _dataService.ExecuteQueryAsync(queryRequest!);

                        response.Result = new
                        {
                            content = new[]
                            {
                                new
                                {
                                    type = "text",
                                    text = JsonSerializer.Serialize(queryResponse, _jsonOptions)
                                }
                            }
                        };
                    }
                    else if (toolName == "get_statistics")
                    {
                        var stats = new
                        {
                            totalRecords = _dataService.GetRecordCount(),
                            lastLoadTime = _dataService.GetLastLoadTime(),
                            dataLoaded = _dataService.GetRecordCount() > 0
                        };

                        response.Result = new
                        {
                            content = new[]
                            {
                                new
                                {
                                    type = "text",
                                    text = JsonSerializer.Serialize(stats, _jsonOptions)
                                }
                            }
                        };
                    }
                    break;

                default:
                    response.Error = new McpError
                    {
                        Code = -32601,
                        Message = $"Method not found: {request.Method}"
                    };
                    break;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error handling MCP request");
            response.Error = new McpError
            {
                Code = -32603,
                Message = ex.Message
            };
        }

        return response;
    }
}

// MCP Protocol Models
public class McpRequest
{
    public string Id { get; set; } = string.Empty;
    public string Method { get; set; } = string.Empty;
    public JsonElement? Params { get; set; }
}

public class McpResponse
{
    public string Id { get; set; } = string.Empty;
    public object? Result { get; set; }
    public McpError? Error { get; set; }
}

public class McpError
{
    public int Code { get; set; }
    public string Message { get; set; } = string.Empty;
}