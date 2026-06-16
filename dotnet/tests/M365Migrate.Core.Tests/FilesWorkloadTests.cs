using System.Net;
using System.Text;
using System.Text.Json;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class FilesWorkloadTests
{
    private const string Src = "/users/jane@contoso.onmicrosoft.com/drive";
    private const string Tgt = "/users/jane@fabrikam.onmicrosoft.com/drive";

    private static JsonElement Parse(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    [Fact]
    public void FromGraph_ParsesPathAndGrants()
    {
        var data = Parse(
            "{\"id\":\"f1\",\"name\":\"Report.txt\",\"size\":10," +
            "\"parentReference\":{\"path\":\"/drive/root:/Reports/2025\"}," +
            "\"permissions\":[{\"roles\":[\"write\"],\"grantedToV2\":{\"user\":{\"userPrincipalName\":\"bob@contoso.onmicrosoft.com\"}}}]}");
        var item = DriveItem.FromGraph(data);

        Assert.Equal("Reports/2025", item.ParentPath);
        Assert.Equal("Reports/2025/Report.txt", item.RelativePath);
        Assert.False(item.IsFolder);
        Assert.Equal("bob@contoso.onmicrosoft.com", item.Grants[0].Upn);
        Assert.Contains("write", item.Grants[0].Roles);
    }

    [Fact]
    public void ResolveDriveRoots_UserAndSite()
    {
        var workload = new FilesWorkload(TestData.Config());
        var (s, t) = workload.ResolveDriveRoots(user: "jane@contoso.onmicrosoft.com");
        Assert.Equal(Src, s);
        Assert.Equal(Tgt, t);

        var (ss, ts) = workload.ResolveDriveRoots(site: "s1", targetSite: "s2");
        Assert.Equal("/sites/s1/drive", ss);
        Assert.Equal("/sites/s2/drive", ts);

        Assert.Throws<ArgumentException>(() => workload.ResolveDriveRoots());
        Assert.Throws<ArgumentException>(() => workload.ResolveDriveRoots(site: "s1"));
    }

    [Fact]
    public void Plan_RewritesGrantsAndMarksCopy()
    {
        var workload = new FilesWorkload(TestData.Config());
        var items = new[]
        {
            new DriveItem { Id = "f1", Name = "ok.txt", Size = 10,
                Grants = { new DriveGrant { Upn = "bob@contoso.onmicrosoft.com", Roles = { "read" } } } },
            new DriveItem { Id = "big", Name = "big.bin", Size = FilesWorkload.SimpleUploadLimit + 1 },
        };
        var planned = workload.Plan(items);

        Assert.All(planned, p => Assert.Equal("copy", p.Action));
        Assert.Equal("bob@fabrikam.onmicrosoft.com", planned[0].TargetGrants[0].Upn);
    }

    [Fact]
    public async Task Discover_WalksBreadthFirst()
    {
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (uri.Contains("/root/children"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"d1\",\"name\":\"Reports\",\"folder\":{},\"parentReference\":{\"path\":\"/drive/root:\"}}," +
                    "{\"id\":\"f0\",\"name\":\"top.txt\",\"parentReference\":{\"path\":\"/drive/root:\"}}]}");
            if (uri.Contains("/items/d1/children"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"f1\",\"name\":\"inner.txt\",\"parentReference\":{\"path\":\"/drive/root:/Reports\"}}]}");
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new FilesWorkload(TestData.Config());

        var items = await workload.DiscoverDriveItemsAsync(client, Src);
        Assert.Equal(new[] { "Reports", "top.txt", "Reports/inner.txt" }, items.Select(i => i.RelativePath));
    }

    [Fact]
    public async Task Migrate_Execute_UploadsSmallFileAndGrants()
    {
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (req.Method == HttpMethod.Get && uri.Contains("/users") && uri.Contains("$select=id"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"tb\",\"userPrincipalName\":\"bob@fabrikam.onmicrosoft.com\"}]}");
            if (req.Method == HttpMethod.Get && uri.Contains("/content"))
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(Encoding.UTF8.GetBytes("hello bytes")) };
            if (req.Method == HttpMethod.Put && uri.Contains("/content"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"tf1\"}");
            if (req.Method == HttpMethod.Post && uri.Contains("/invite"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new FilesWorkload(TestData.Config());

        var planned = new[]
        {
            new PlannedDriveItem
            {
                SourceId = "f1", Name = "ok.txt", RelativePath = "ok.txt", IsFolder = false, Size = 11, Action = "copy",
                TargetGrants = { new DriveGrant { Upn = "bob@fabrikam.onmicrosoft.com", Roles = { "write" } } },
            },
        };
        var results = await workload.MigrateDriveItemsAsync(client, client, planned, Src, Tgt, dryRun: false);

        Assert.Equal("uploaded", results[0].Status);
        Assert.Equal("granted:1", results[0].Detail["grants"]);
        Assert.Contains("hello bytes", handler.Bodies);
    }

    [Fact]
    public async Task Migrate_Execute_UploadsLargeFileViaSession()
    {
        var sessionCreated = false;
        var chunkPut = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (req.Method == HttpMethod.Get && uri.Contains("/users") && uri.Contains("$select=id"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
            if (req.Method == HttpMethod.Get && uri.Contains("/content"))
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(Encoding.UTF8.GetBytes("large payload")) };
            if (req.Method == HttpMethod.Post && uri.Contains("createUploadSession"))
            {
                sessionCreated = true;
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"uploadUrl\":\"https://upload.example/sess\"}");
            }
            if (req.Method == HttpMethod.Put && uri.Contains("upload.example"))
            {
                chunkPut = true;
                return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"tbig\"}");
            }
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new FilesWorkload(TestData.Config());

        var planned = new[]
        {
            new PlannedDriveItem
            {
                SourceId = "big", Name = "big.bin", RelativePath = "big.bin",
                IsFolder = false, Size = FilesWorkload.SimpleUploadLimit + 1, Action = "copy",
            },
        };
        var results = await workload.MigrateDriveItemsAsync(client, client, planned, Src, Tgt, dryRun: false);

        Assert.True(sessionCreated && chunkPut);
        Assert.Equal("uploaded-session", results[0].Status);
    }
}
