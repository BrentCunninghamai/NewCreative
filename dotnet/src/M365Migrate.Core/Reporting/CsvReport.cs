using System.Text;

namespace M365Migrate.Core.Reporting;

/// <summary>Builds RFC 4180-style CSV text for run reports (plans and results).</summary>
public static class CsvReport
{
    public static string ToCsv(IReadOnlyList<string> headers, IEnumerable<IReadOnlyList<string>> rows)
    {
        var sb = new StringBuilder();
        sb.AppendLine(string.Join(",", headers.Select(Escape)));
        foreach (var row in rows)
            sb.AppendLine(string.Join(",", row.Select(Escape)));
        return sb.ToString();
    }

    /// <summary>Quote a field if it contains a comma, quote, or newline; double inner quotes.</summary>
    public static string Escape(string? field)
    {
        field ??= "";
        if (field.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0)
            return "\"" + field.Replace("\"", "\"\"") + "\"";
        return field;
    }
}
