using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class SharePointWorkloadTests
{
    [Theory]
    [InlineData("https://contoso.sharepoint.com/sites/Marketing", "contoso.sharepoint.com", "/sites/Marketing")]
    [InlineData("https://contoso.sharepoint.com/sites/Marketing/", "contoso.sharepoint.com", "/sites/Marketing")]
    [InlineData("https://contoso.sharepoint.com", "contoso.sharepoint.com", "")]
    public void ParseSiteUrl_SplitsHostAndPath(string url, string host, string path)
    {
        var (h, p) = SharePointWorkload.ParseSiteUrl(url);
        Assert.Equal(host, h);
        Assert.Equal(path, p);
    }

    [Fact]
    public void MatchDrives_PairsByNameAndReportsUnmatched()
    {
        var source = new[]
        {
            new SiteDrive("s1", "Documents"),
            new SiteDrive("s2", "Policies"),
            new SiteDrive("s3", "Archive"),
        };
        var target = new[]
        {
            new SiteDrive("t1", "Documents"),
            new SiteDrive("t2", "Policies"),
        };

        var (pairs, unmatched) = SharePointWorkload.MatchDrives(source, target);

        Assert.Equal(2, pairs.Count);
        Assert.Equal("t1", pairs.Single(p => p.Source.Name == "Documents").Target.Id);
        Assert.Equal(new[] { "Archive" }, unmatched);
    }

    [Fact]
    public void DriveRoot_FormatsForFilesEngine()
    {
        Assert.Equal("/drives/abc", SharePointWorkload.DriveRoot("abc"));
    }
}
