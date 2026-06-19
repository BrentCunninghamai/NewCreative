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

    [Theory]
    [InlineData("https://contoso.sharepoint.com", "contoso-my.sharepoint.com")]
    [InlineData("https://net12.sharepoint.com/", "net12-my.sharepoint.com")]
    [InlineData("https://contoso.sharepoint.us", "contoso-my.sharepoint.us")]
    public void MyHostFromRoot_InsertsMy(string rootUrl, string expected)
    {
        Assert.Equal(expected, SharePointWorkload.MyHostFromRoot(rootUrl));
    }

    [Fact]
    public void OneDriveUrl_MungesUpnIntoPersonalSite()
    {
        Assert.Equal(
            "https://net12-my.sharepoint.com/personal/melicia_pieterse_net1_com",
            SharePointWorkload.OneDriveUrl("net12-my.sharepoint.com", "melicia.pieterse@net1.com"));
        // Non-alphanumeric characters (incl. apostrophes/hyphens) all become underscores.
        Assert.Equal(
            "https://h-my.sharepoint.com/personal/o_brien_jp_x_co_za",
            SharePointWorkload.OneDriveUrl("h-my.sharepoint.com", "o'brien-jp@x.co.za"));
    }
}
