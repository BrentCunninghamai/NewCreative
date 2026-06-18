using System.Net;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class TeamsMessagesWorkloadTests
{
    private static MigratableTeam SampleTeam() => new()
    {
        GroupId = "sg1",
        DisplayName = "Sales",
        MailNickname = "sales",
        CreatedDateTime = "2021-01-01T00:00:00Z",
        Channels = { new MigratableChannel { Id = "src-general", DisplayName = "General" } },
    };

    [Fact]
    public async Task Migrate_CreatesMigrationTeam_ImportsMessage_CompletesMigration()
    {
        var completeCalled = false;
        string? messageBody = null;
        var handler = new FakeHttpMessageHandler((req, body) =>
        {
            var uri = req.RequestUri!.ToString();

            if (req.Method == HttpMethod.Get && uri.Contains("/users"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"su1\",\"userPrincipalName\":\"alice@contoso.onmicrosoft.com\"}," +
                    "{\"id\":\"tu1\",\"userPrincipalName\":\"alice@fabrikam.onmicrosoft.com\"}]}");

            if (req.Method == HttpMethod.Post && uri.EndsWith("/teams"))
            {
                var r = new HttpResponseMessage(HttpStatusCode.Accepted);
                r.Headers.Location = new Uri("/teams('19:newteam@thread.tacv2')", UriKind.Relative);
                return r;
            }
            if (req.Method == HttpMethod.Post && uri.Contains("completeMigration"))
            {
                completeCalled = true;
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            }
            if (req.Method == HttpMethod.Post && uri.Contains("/messages"))
            {
                messageBody = body;
                return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"m1\"}");
            }
            // GET source channel messages
            if (req.Method == HttpMethod.Get && uri.Contains("/teams/sg1/channels/") && uri.Contains("/messages"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"messageType\":\"message\",\"createdDateTime\":\"2021-05-05T10:00:00Z\"," +
                    "\"from\":{\"user\":{\"id\":\"su1\",\"displayName\":\"Alice\"}}," +
                    "\"body\":{\"contentType\":\"html\",\"content\":\"<p>hi</p>\"}}]}");
            // GET target team channels (General auto-created)
            if (req.Method == HttpMethod.Get && uri.Contains("19:newteam") && uri.Contains("/channels"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"ch-general\",\"displayName\":\"General\"}]}");

            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var results = await new TeamsMessagesWorkload(TestData.Config())
            .MigrateAsync(client, client, new[] { SampleTeam() }, dryRun: false);

        Assert.Equal("ok", results[0].Status);
        Assert.Equal("imported:1", results[0].Detail["messages"]);
        Assert.True(completeCalled);
        // The imported message carries the resolved target author id.
        Assert.NotNull(messageBody);
        Assert.Contains("tu1", messageBody!);
    }

    [Fact]
    public async Task Migrate_DryRun_WritesNothing()
    {
        var posted = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            if (req.Method == HttpMethod.Post) posted = true;
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var results = await new TeamsMessagesWorkload(TestData.Config())
            .MigrateAsync(client, client, new[] { SampleTeam() }, dryRun: true);

        Assert.False(posted);
        Assert.Equal("would-migrate", results[0].Status);
        Assert.Equal("1", results[0].Detail["channels"]);
    }

    [Fact]
    public async Task Migrate_SkipsMessageWithUnresolvableAuthor()
    {
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (req.Method == HttpMethod.Get && uri.Contains("/users"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}"); // no users -> no author resolves
            if (req.Method == HttpMethod.Post && uri.EndsWith("/teams"))
            {
                var r = new HttpResponseMessage(HttpStatusCode.Accepted);
                r.Headers.Location = new Uri("/teams('19:newteam@thread.tacv2')", UriKind.Relative);
                return r;
            }
            if (req.Method == HttpMethod.Post && uri.Contains("completeMigration"))
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            if (req.Method == HttpMethod.Get && uri.Contains("/teams/sg1/channels/") && uri.Contains("/messages"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"messageType\":\"message\",\"from\":{\"user\":{\"id\":\"ghost\"}},\"body\":{\"content\":\"x\"}}]}");
            if (req.Method == HttpMethod.Get && uri.Contains("19:newteam") && uri.Contains("/channels"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[{\"id\":\"ch-general\",\"displayName\":\"General\"}]}");
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var results = await new TeamsMessagesWorkload(TestData.Config())
            .MigrateAsync(client, client, new[] { SampleTeam() }, dryRun: false);

        Assert.Equal("ok", results[0].Status);
        Assert.Equal("imported:0", results[0].Detail["messages"]);
        Assert.Equal("1", results[0].Detail["skipped"]);
    }

    [Fact]
    public void MigrationMessageBody_UsesAuthorContentAndTimestamp()
    {
        var msg = new ChannelMessage
        {
            CreatedDateTime = "2021-05-01T10:00:00Z",
            FromDisplayName = "Jane Doe",
            BodyContentType = "html",
            BodyContent = "<p>hello</p>",
        };

        var body = TeamsMessagesWorkload.MigrationMessageBody("target-id-1", msg, "2020-01-01T00:00:00Z");

        Assert.Equal("2021-05-01T10:00:00Z", body["createdDateTime"]);
        var from = Assert.IsAssignableFrom<Dictionary<string, object>>(body["from"]);
        var user = Assert.IsAssignableFrom<Dictionary<string, object>>(from["user"]);
        Assert.Equal("target-id-1", user["id"]);
        Assert.Equal("aadUser", user["userIdentityType"]);
        var bodyEl = Assert.IsAssignableFrom<Dictionary<string, object>>(body["body"]);
        Assert.Equal("<p>hello</p>", bodyEl["content"]);
    }

    [Fact]
    public void MigrationMessageBody_FallsBackWhenNoTimestamp()
    {
        var msg = new ChannelMessage { CreatedDateTime = null, BodyContent = "x" };
        var body = TeamsMessagesWorkload.MigrationMessageBody("id", msg, "2020-01-01T00:00:00Z");
        Assert.Equal("2020-01-01T00:00:00Z", body["createdDateTime"]);
    }
}
