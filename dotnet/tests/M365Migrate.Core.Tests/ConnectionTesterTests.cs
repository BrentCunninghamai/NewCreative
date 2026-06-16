using System.Net;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class ConnectionTesterTests
{
    [Fact]
    public async Task Check_ReturnsOrganizationDisplayName()
    {
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            Assert.Contains("/organization", req.RequestUri!.ToString());
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[{\"displayName\":\"Fabrikam Inc\"}]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var org = await ConnectionTester.CheckAsync(client);
        Assert.Equal("Fabrikam Inc", org);
    }

    [Fact]
    public async Task Check_PropagatesGraphErrorOnAuthFailure()
    {
        var handler = new FakeHttpMessageHandler((_, __) =>
            FakeHttpMessageHandler.Text(HttpStatusCode.Unauthorized, "invalid client secret"));
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var ex = await Assert.ThrowsAsync<GraphException>(() => ConnectionTester.CheckAsync(client));
        Assert.Equal(401, ex.StatusCode);
    }
}
