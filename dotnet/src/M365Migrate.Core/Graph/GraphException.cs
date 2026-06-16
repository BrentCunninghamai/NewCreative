namespace M365Migrate.Core.Graph;

/// <summary>Raised for non-retryable / exhausted Graph API errors.</summary>
public sealed class GraphException : Exception
{
    public int StatusCode { get; }

    public GraphException(int statusCode, string message)
        : base($"Graph API error {statusCode}: {message}")
    {
        StatusCode = statusCode;
    }
}
