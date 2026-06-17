using System.Net;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class CalendarContactsWorkloadTests
{
    private const string Src = "/users/jane@contoso.onmicrosoft.com";
    private const string Tgt = "/users/jane@fabrikam.onmicrosoft.com";

    [Fact]
    public void ResolveUserRefs_UsesRewrittenTargetUpn()
    {
        var (s, t) = new CalendarContactsWorkload(TestData.Config()).ResolveUserRefs("jane@contoso.onmicrosoft.com");
        Assert.Equal(Src, s);
        Assert.Equal(Tgt, t);
    }

    [Fact]
    public async Task Migrate_CopiesNewItemsAndSkipsExisting()
    {
        var eventPosts = 0;
        var contactPosts = 0;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            var isTarget = uri.Contains("fabrikam");

            if (req.Method == HttpMethod.Get && uri.Contains("/events"))
            {
                if (isTarget)
                    // target already has "Standup"
                    return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                        "{\"value\":[{\"subject\":\"Standup\",\"start\":{\"dateTime\":\"2026-01-01T09:00:00\"},\"end\":{\"dateTime\":\"2026-01-01T09:15:00\"}}]}");
                // source has the duplicate Standup + a new Planning event
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[" +
                    "{\"subject\":\"Standup\",\"start\":{\"dateTime\":\"2026-01-01T09:00:00\"},\"end\":{\"dateTime\":\"2026-01-01T09:15:00\"}}," +
                    "{\"subject\":\"Planning\",\"start\":{\"dateTime\":\"2026-01-02T10:00:00\"},\"end\":{\"dateTime\":\"2026-01-02T11:00:00\"}}]}");
            }
            if (req.Method == HttpMethod.Get && uri.Contains("/contacts"))
            {
                if (isTarget)
                    return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"displayName\":\"Bob\",\"emailAddresses\":[{\"address\":\"bob@contoso.com\"}]}]}");
            }
            if (req.Method == HttpMethod.Post && uri.Contains("/events")) { eventPosts++; return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"e\"}"); }
            if (req.Method == HttpMethod.Post && uri.Contains("/contacts")) { contactPosts++; return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"c\"}"); }
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var results = await new CalendarContactsWorkload(TestData.Config())
            .MigrateAsync(client, client, Src, Tgt, dryRun: false);

        var cal = results.Single(r => r.Name == "Calendar");
        var con = results.Single(r => r.Name == "Contacts");
        Assert.Equal("1", cal.Detail["copied"]);   // Planning copied
        Assert.Equal("1", cal.Detail["skipped"]);   // Standup skipped (dup)
        Assert.Equal("1", con.Detail["copied"]);    // Bob copied
        Assert.Equal(1, eventPosts);
        Assert.Equal(1, contactPosts);
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

        var results = await new CalendarContactsWorkload(TestData.Config())
            .MigrateAsync(client, client, Src, Tgt, dryRun: true);

        Assert.False(posted);
        Assert.All(results, r => Assert.Equal("would-copy", r.Status));
    }
}
