using ClosedXML.Excel;
using ExcelMcpServer.Models;
using Microsoft.Extensions.Logging;
using System.Collections.Concurrent;
using System.Diagnostics;

namespace ExcelMcpServer.Services;

public class ExcelDataService
{
    private readonly ILogger<ExcelDataService> _logger;
    private List<LoanRecord> _cachedData = new();
    private readonly object _lockObj = new();
    private DateTime _lastLoadTime;

    // Indexed collections for fast lookups
    private Dictionary<string, LoanRecord> _loanIndex = new();
    private List<LoanRecord> _mismatchCache = new();

    public ExcelDataService(ILogger<ExcelDataService> logger)
    {
        _logger = logger;
    }

    public async Task<bool> LoadExcelFileAsync(string filePath)
    {
        var sw = Stopwatch.StartNew();

        try
        {
            _logger.LogInformation("Starting to load Excel file: {FilePath}", filePath);

            var records = await Task.Run(() => LoadExcelData(filePath));

            lock (_lockObj)
            {
                _cachedData = records;
                BuildIndexes();
                _lastLoadTime = DateTime.UtcNow;
            }

            sw.Stop();
            _logger.LogInformation(
                "Successfully loaded {Count} records in {ElapsedMs}ms",
                _cachedData.Count,
                sw.ElapsedMilliseconds);

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading Excel file");
            return false;
        }
    }

    private List<LoanRecord> LoadExcelData(string filePath)
    {
        var records = new List<LoanRecord>(85000); // Pre-allocate for performance

        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheet(1);
        var rows = worksheet.RowsUsed().Skip(1); // Skip header

        foreach (var row in rows)
        {
            try
            {
                var record = new LoanRecord
                {
                    LoanID = row.Cell(1).GetValue<string>(),
                    BorrowerName = row.Cell(2).GetValue<string>(),
                    Servicer_LoanAmount = row.Cell(3).GetValue<decimal>(),
                    FNMA_LoanAmount = row.Cell(4).GetValue<decimal>(),
                    DifferenceAmount = row.Cell(5).GetValue<decimal>(),
                    ReconciledStatus = row.Cell(6).GetValue<string>()
                };

                records.Add(record);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error parsing row {RowNumber}", row.RowNumber());
            }
        }

        return records;
    }

    private void BuildIndexes()
    {
        // Build loan ID index for O(1) lookups
        _loanIndex = _cachedData.ToDictionary(r => r.LoanID);

        // Cache mismatches for frequent queries
        _mismatchCache = _cachedData.Where(r => r.HasMismatch).ToList();

        _logger.LogInformation(
            "Built indexes: {TotalRecords} total, {MismatchCount} mismatches",
            _cachedData.Count,
            _mismatchCache.Count);
    }

    public async Task<QueryResponse> ExecuteQueryAsync(QueryRequest request)
    {
        var sw = Stopwatch.StartNew();
        var response = new QueryResponse { Success = true };

        try
        {
            var query = request.Query.ToLowerInvariant();
            IEnumerable<LoanRecord> results;
            string queryType;

            // Enhanced natural language query parsing
            if (query.Contains("mismatch") || query.Contains("reconcile") ||
                (query.Contains("difference") && !query.Contains("where") && !query.Contains(">")))
            {
                queryType = "FindMismatches";
                results = _mismatchCache;
            }
            else if (query.Contains("difference") && (query.Contains(">") || query.Contains("greater")))
            {
                queryType = "DifferenceGreaterThan";
                var threshold = ExtractNumber(query);
                results = _cachedData.Where(r => r.DifferenceAmount > threshold);
            }
            else if (query.Contains("difference") && (query.Contains("<") || query.Contains("less")))
            {
                queryType = "DifferenceLessThan";
                var threshold = ExtractNumber(query);
                results = _cachedData.Where(r => r.DifferenceAmount < threshold);
            }
            else if (query.Contains("reconciled") && !query.Contains("un") && !query.Contains("not"))
            {
                queryType = "ReconciledLoans";
                results = _cachedData.Where(r =>
                    r.ReconciledStatus.Equals("reconciled", StringComparison.OrdinalIgnoreCase));
            }
            else if (query.Contains("unreconciled") || query.Contains("not reconciled") || query.Contains("pending"))
            {
                queryType = "UnreconciledLoans";
                results = _cachedData.Where(r =>
                    !r.ReconciledStatus.Equals("reconciled", StringComparison.OrdinalIgnoreCase));
            }
            else if ((query.Contains("loan") && (query.Contains("id") || query.Contains("number"))) ||
                     System.Text.RegularExpressions.Regex.IsMatch(query, @"\bLN-?\d+\b",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                queryType = "LoanByID";
                var loanId = ExtractLoanId(query);
                results = _loanIndex.ContainsKey(loanId)
                    ? new[] { _loanIndex[loanId] }
                    : Array.Empty<LoanRecord>();
            }
            else if (query.Contains("borrower") || query.Contains("customer") || query.Contains("name"))
            {
                queryType = "SearchByBorrower";
                var borrowerName = ExtractBorrowerName(query);
                if (!string.IsNullOrEmpty(borrowerName))
                {
                    results = _cachedData.Where(r =>
                        r.BorrowerName.Contains(borrowerName, StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    results = Array.Empty<LoanRecord>();
                    response.Message = "Please specify a borrower name to search for.";
                }
            }
            else if (query.Contains("top") || query.Contains("highest") || query.Contains("largest"))
            {
                queryType = "TopDifferences";
                var count = ExtractNumber(query);
                if (count == 0) count = 10;
                results = _cachedData
                    .OrderByDescending(r => Math.Abs(r.DifferenceAmount))
                    .Take((int)count);
            }
            else if (query.Contains("bottom") || query.Contains("lowest") || query.Contains("smallest"))
            {
                queryType = "BottomDifferences";
                var count = ExtractNumber(query);
                if (count == 0) count = 10;
                results = _cachedData
                    .Where(r => r.DifferenceAmount != 0)
                    .OrderBy(r => Math.Abs(r.DifferenceAmount))
                    .Take((int)count);
            }
            else if (query.Contains("positive") && query.Contains("difference"))
            {
                queryType = "PositiveDifferences";
                results = _cachedData.Where(r => r.DifferenceAmount > 0);
            }
            else if (query.Contains("negative") && query.Contains("difference"))
            {
                queryType = "NegativeDifferences";
                results = _cachedData.Where(r => r.DifferenceAmount < 0);
            }
            else if (query.Contains("servicer") && query.Contains("greater") ||
                     query.Contains("servicer") && query.Contains("more"))
            {
                queryType = "ServicerGreaterThanFNMA";
                results = _cachedData.Where(r => r.Servicer_LoanAmount > r.FNMA_LoanAmount);
            }
            else if (query.Contains("fnma") && query.Contains("greater") ||
                     query.Contains("fnma") && query.Contains("more"))
            {
                queryType = "FNMAGreaterThanServicer";
                results = _cachedData.Where(r => r.FNMA_LoanAmount > r.Servicer_LoanAmount);
            }
            else if (query.Contains("count") || query.Contains("how many") || query.Contains("total"))
            {
                queryType = "Count";
                results = _cachedData;
                response.Message = $"Total count: {_cachedData.Count:N0} records";
            }
            else if (query.Contains("all") || query.Contains("list") || query.Contains("show") || query.Contains("everything"))
            {
                queryType = "ListAll";
                results = _cachedData;
            }
            else if (query.Contains("summary") || query.Contains("overview") || query.Contains("report"))
            {
                queryType = "Summary";
                results = _cachedData.Take(10); // Sample
                response.Message = GetSummaryMessage();
            }
            else
            {
                queryType = "Unknown";
                results = Array.Empty<LoanRecord>();
                response.Message = "I didn't understand that query. Try:\n" +
                    "• 'Find mismatches'\n" +
                    "• 'Show loans where difference > 5000'\n" +
                    "• 'List unreconciled loans'\n" +
                    "• 'Find loan LN-12345'\n" +
                    "• 'Search borrower John Smith'\n" +
                    "Type 'help' for more examples.";
            }

            var resultList = results.ToList();
            response.TotalCount = resultList.Count;

            // Apply pagination
            var skip = request.Skip ?? 0;
            var limit = request.Limit ?? 100;
            response.Data = resultList.Skip(skip).Take(limit).ToList();

            // Build metadata
            response.Metadata = new QueryMetadata
            {
                QueryType = queryType,
                ExecutionTimeMs = sw.ElapsedMilliseconds,
                Statistics = BuildStatistics(resultList)
            };

            if (string.IsNullOrEmpty(response.Message))
            {
                response.Message = $"Found {response.TotalCount:N0} records matching query";
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error executing query");
            response.Success = false;
            response.Message = $"Query execution failed: {ex.Message}";
        }

        return response;
    }

    private string GetSummaryMessage()
    {
        var total = _cachedData.Count;
        var mismatches = _mismatchCache.Count;
        var reconciled = _cachedData.Count(r =>
            r.ReconciledStatus.Equals("reconciled", StringComparison.OrdinalIgnoreCase));

        return $"Dataset Summary: {total:N0} total loans, {mismatches:N0} mismatches ({(mismatches * 100.0 / total):F1}%), {reconciled:N0} reconciled";
    }

    private Dictionary<string, object> BuildStatistics(List<LoanRecord> data)
    {
        if (!data.Any()) return new Dictionary<string, object>();

        return new Dictionary<string, object>
        {
            ["TotalAmount_Servicer"] = data.Sum(r => r.Servicer_LoanAmount),
            ["TotalAmount_FNMA"] = data.Sum(r => r.FNMA_LoanAmount),
            ["TotalDifference"] = data.Sum(r => r.DifferenceAmount),
            ["AverageDifference"] = data.Average(r => r.DifferenceAmount),
            ["MaxDifference"] = data.Max(r => r.DifferenceAmount),
            ["MinDifference"] = data.Min(r => r.DifferenceAmount),
            ["MismatchCount"] = data.Count(r => r.HasMismatch)
        };
    }

    private decimal ExtractNumber(string query)
    {
        var match = System.Text.RegularExpressions.Regex.Match(query, @"\d+\.?\d*");
        return match.Success ? decimal.Parse(match.Value) : 0;
    }

    private string ExtractLoanId(string query)
    {
        var match = System.Text.RegularExpressions.Regex.Match(query, @"[A-Z0-9\-]+",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return match.Success ? match.Value : string.Empty;
    }

    private string ExtractBorrowerName(string query)
    {
        var parts = query.Split(new[] { "borrower", "name" }, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length > 1 ? parts[1].Trim() : string.Empty;
    }

    public int GetRecordCount() => _cachedData.Count;
    public DateTime GetLastLoadTime() => _lastLoadTime;
}