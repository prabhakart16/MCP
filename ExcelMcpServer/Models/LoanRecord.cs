namespace ExcelMcpServer.Models;

public class LoanRecord
{
    public string LoanID { get; set; } = string.Empty;
    public string BorrowerName { get; set; } = string.Empty;
    public decimal Servicer_LoanAmount { get; set; }
    public decimal FNMA_LoanAmount { get; set; }
    public decimal DifferenceAmount { get; set; }
    public string ReconciledStatus { get; set; } = string.Empty;

    // Computed property for quick mismatch detection
    public bool HasMismatch => DifferenceAmount != 0;
}

public class QueryRequest
{
    public string Query { get; set; } = string.Empty;
    public int? Limit { get; set; }
    public int? Skip { get; set; }
}

public class QueryResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<LoanRecord> Data { get; set; } = new();
    public int TotalCount { get; set; }
    public QueryMetadata Metadata { get; set; } = new();
}

public class QueryMetadata
{
    public string QueryType { get; set; } = string.Empty;
    public long ExecutionTimeMs { get; set; }
    public Dictionary<string, object> Statistics { get; set; } = new();
}