using M365Migrate.Core.Reporting;
using Xunit;

namespace M365Migrate.Core.Tests;

public class CsvReportTests
{
    [Fact]
    public void Escape_QuotesFieldsWithSpecialCharacters()
    {
        Assert.Equal("plain", CsvReport.Escape("plain"));
        Assert.Equal("\"a,b\"", CsvReport.Escape("a,b"));
        Assert.Equal("\"says \"\"hi\"\"\"", CsvReport.Escape("says \"hi\""));
        Assert.Equal("\"line1\nline2\"", CsvReport.Escape("line1\nline2"));
    }

    [Fact]
    public void ToCsv_WritesHeaderAndRows()
    {
        var csv = CsvReport.ToCsv(
            new[] { "Name", "Status" },
            new[]
            {
                new[] { "jane@fabrikam.com", "created" },
                new[] { "bob, jr", "skipped" },
            });

        var lines = csv.TrimEnd().Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
        Assert.Equal("Name,Status", lines[0]);
        Assert.Equal("jane@fabrikam.com,created", lines[1]);
        Assert.Equal("\"bob, jr\",skipped", lines[2]);
    }
}
