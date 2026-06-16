using System.Net;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using Xunit;

namespace M365Migrate.Core.Tests;

public class GraphClientTests
{
    private static GraphClient Build(FakeHttpMessageHandler handler, Action<int>? onDelay = null) =>
        new(new HttpClient(handler), TokenProviders.Static("fake-token"),
            delay: ms => { onDelay?.Invoke(ms); return Task.CompletedTask; });

    [Fact]
    public async Task GetAll_FollowsPagination()
    {
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            return uri.Contains("skiptoken")
                ? FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[{\"id\":\"3\"}]}")
                : FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"1\"},{\"id\":\"2\"}],\"@odata.nextLink\":\"https://graph.microsoft.com/v1.0/users?$skiptoken=abc\"}");
        });

        var client = Build(handler);
        var items = await client.GetAllAsync("/users");

        var ids = items.Select(i => i.GetStringOrNull("id")).ToList();
        Assert.Equal(new[] { "1", "2", "3" }, ids);
    }

    [Fact]
    public async Task Retries_On429_ThenSucceeds()
    {
        var calls = 0;
        var handler = new FakeHttpMessageHandler((_, __) =>
        {
            calls++;
            if (calls == 1)
            {
                var resp = FakeHttpMessageHandler.Text(HttpStatusCode.TooManyRequests, "throttled");
                resp.Headers.TryAddWithoutValidation("Retry-After", "0");
                return resp;
            }
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[{\"displayName\":\"Fabrikam\"}]}");
        });

        var delays = new List<int>();
        var client = Build(handler, delays.Add);
        var data = await client.GetAsync("/organization");

        Assert.Single(delays);
        Assert.Equal(0, delays[0]); // honored Retry-After: 0
        var org = data.GetProperty("value")[0].GetStringOrNull("displayName");
        Assert.Equal("Fabrikam", org);
    }

    [Fact]
    public async Task Throws_GraphException_On4xx()
    {
        var handler = new FakeHttpMessageHandler((_, __) =>
            FakeHttpMessageHandler.Text(HttpStatusCode.Forbidden, "forbidden"));
        var client = Build(handler);

        var ex = await Assert.ThrowsAsync<GraphException>(() => client.GetAsync("/users"));
        Assert.Equal(403, ex.StatusCode);
    }
}
