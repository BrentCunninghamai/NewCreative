using System.Net;
using System.Text;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class MailWorkloadTests
{
    private const string Src = "/users/jane@contoso.onmicrosoft.com";
    private const string Tgt = "/users/jane@fabrikam.onmicrosoft.com";

    [Fact]
    public void ResolveUserRefs_UsesRewrittenTargetUpn()
    {
        var workload = new MailWorkload(TestData.Config());
        var (s, t) = workload.ResolveUserRefs("jane@contoso.onmicrosoft.com");
        Assert.Equal(Src, s);
        Assert.Equal(Tgt, t);
        Assert.Throws<ArgumentException>(() => workload.ResolveUserRefs(""));
    }

    [Fact]
    public async Task Discover_WalksFolderTreeBreadthFirst()
    {
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (uri.Contains("/mailFolders/s-inbox/childFolders"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"s-2025\",\"displayName\":\"2025\",\"totalItemCount\":3,\"childFolderCount\":0}]}");
            if (uri.Contains("/mailFolders?"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"s-inbox\",\"displayName\":\"Inbox\",\"totalItemCount\":10,\"childFolderCount\":1}," +
                    "{\"id\":\"s-arch\",\"displayName\":\"Archive\",\"totalItemCount\":5,\"childFolderCount\":0}]}");
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var folders = await new MailWorkload(TestData.Config()).DiscoverFoldersAsync(client, Src);
        Assert.Equal(new[] { "Inbox", "Archive", "Inbox/2025" }, folders.Select(f => f.Path));
        Assert.Equal(10, folders[0].TotalItemCount);
    }

    [Fact]
    public async Task Migrate_CopiesNewMessageAndSkipsDuplicate()
    {
        var dupId = "<dup@contoso>";
        var newId = "<new@contoso>";
        byte[]? postedBody = null;
        var handler = new FakeHttpMessageHandler((req, body) =>
        {
            var uri = req.RequestUri!.ToString();
            // target folder discovery: Inbox already exists
            if (req.Method == HttpMethod.Get && uri.Contains("fabrikam") && uri.Contains("/mailFolders?"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"t-inbox\",\"displayName\":\"Inbox\",\"totalItemCount\":1,\"childFolderCount\":0}]}");
            // existing target messages, scanned mailbox-wide (idempotency seed)
            if (req.Method == HttpMethod.Get && uri.Contains("fabrikam") && uri.Contains("internetMessageId"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    $"{{\"value\":[{{\"internetMessageId\":\"{dupId}\"}}]}}");
            // source messages: one duplicate, one new
            if (req.Method == HttpMethod.Get && uri.Contains("/mailFolders/s-inbox/messages"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    $"{{\"value\":[{{\"id\":\"m1\",\"internetMessageId\":\"{dupId}\"}},{{\"id\":\"m2\",\"internetMessageId\":\"{newId}\"}}]}}");
            // MIME of the new message
            if (req.Method == HttpMethod.Get && uri.Contains("/messages/m2/$value"))
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(Encoding.ASCII.GetBytes("RAWMIME")) };
            // create message in target folder
            if (req.Method == HttpMethod.Post && uri.Contains("/mailFolders/t-inbox/messages"))
            {
                postedBody = Encoding.ASCII.GetBytes(body);
                return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"t-m2\"}");
            }
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new MailWorkload(TestData.Config());

        var planned = new[]
        {
            new PlannedMailFolder { SourceId = "s-inbox", DisplayName = "Inbox", Path = "Inbox", ItemCount = 2 },
        };
        var results = await workload.MigrateAsync(client, client, Src, Tgt, planned, dryRun: false);

        Assert.Equal("ok", results[0].Status);
        Assert.Equal("1", results[0].Detail["copied"]);
        Assert.Equal("1", results[0].Detail["skipped"]);
        // The created message body is the base64 of the source MIME.
        Assert.Equal(Convert.ToBase64String(Encoding.ASCII.GetBytes("RAWMIME")),
            Encoding.ASCII.GetString(postedBody!));
    }

    [Fact]
    public async Task Migrate_SkipsMessageAlreadyInADifferentTargetFolder()
    {
        // A message that another tool already placed in the target's Archive must be
        // recognised and skipped even though the source has it in Inbox. This proves
        // the dedup is mailbox-wide, not per-folder.
        var movedId = "<moved@contoso>";
        var postedToTarget = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            // target folder discovery: Inbox exists (Archive is irrelevant to placement)
            if (req.Method == HttpMethod.Get && uri.Contains("fabrikam") && uri.Contains("/mailFolders?"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"t-inbox\",\"displayName\":\"Inbox\",\"totalItemCount\":1,\"childFolderCount\":0}]}");
            // mailbox-wide scan returns the message (it's filed in Archive in the target)
            if (req.Method == HttpMethod.Get && uri.Contains("fabrikam") && uri.Contains("internetMessageId"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    $"{{\"value\":[{{\"internetMessageId\":\"{movedId}\"}}]}}");
            // source Inbox has that same message
            if (req.Method == HttpMethod.Get && uri.Contains("/mailFolders/s-inbox/messages"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    $"{{\"value\":[{{\"id\":\"m1\",\"internetMessageId\":\"{movedId}\"}}]}}");
            if (req.Method == HttpMethod.Post) postedToTarget = true;
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new MailWorkload(TestData.Config());

        var planned = new[]
        {
            new PlannedMailFolder { SourceId = "s-inbox", DisplayName = "Inbox", Path = "Inbox", ItemCount = 1 },
        };
        var results = await workload.MigrateAsync(client, client, Src, Tgt, planned, dryRun: false);

        Assert.Equal("ok", results[0].Status);
        Assert.Equal("0", results[0].Detail["copied"]);
        Assert.Equal("1", results[0].Detail["skipped"]);
        Assert.False(postedToTarget); // nothing re-copied
    }

    [Fact]
    public async Task Migrate_DryRun_WritesNothing()
    {
        var posted = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            if (req.Method == HttpMethod.Post) posted = true;
            if (req.RequestUri!.ToString().Contains("/mailFolders?"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var planned = new[] { new PlannedMailFolder { SourceId = "s-inbox", DisplayName = "Inbox", Path = "Inbox", ItemCount = 5 } };

        var results = await new MailWorkload(TestData.Config()).MigrateAsync(client, client, Src, Tgt, planned, dryRun: true);

        Assert.False(posted);
        Assert.Equal("would-copy", results[0].Status);
    }
}
