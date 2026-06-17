using System.Net;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class UserResolverTests
{
    [Theory]
    [InlineData("/users/paule@net1.com/drive", "paule@net1.com")]
    [InlineData("/users/paule@net1.com", "paule@net1.com")]
    [InlineData("paule@net1.com", "paule@net1.com")]
    [InlineData("  /users/abc-guid  ", "abc-guid")]
    public void NormalizeKey_StripsRefAndTrailingSegment(string input, string expected)
    {
        Assert.Equal(expected, UserResolver.NormalizeKey(input));
    }

    [Fact]
    public async Task TryResolve_ReturnsUser_AndQueriesNormalizedKey()
    {
        string? requested = null;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            requested = req.RequestUri!.ToString();
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                "{\"id\":\"id-1\",\"displayName\":\"Paul Encarnacao\",\"userPrincipalName\":\"PaulE@net1.com\",\"mail\":\"paule@lesakatech.com\"}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var user = await UserResolver.TryResolveAsync(client, "/users/PaulE@net1.com/drive");

        Assert.NotNull(user);
        Assert.Equal("id-1", user!.Id);
        Assert.Equal("Paul Encarnacao (PaulE@net1.com)", user.Describe());
        Assert.Contains("/users/PaulE@net1.com?", requested); // ref + trailing segment stripped
    }

    [Fact]
    public async Task TryResolve_ReturnsNull_On404()
    {
        var handler = new FakeHttpMessageHandler((_, _) =>
            FakeHttpMessageHandler.Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"Request_ResourceNotFound\"}}"));
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        Assert.Null(await UserResolver.TryResolveAsync(client, "ghost@net1.com"));
    }

    [Fact]
    public void Describe_FallsBackToMailThenId()
    {
        Assert.Equal("Jane (jane@x.com)", new ResolvedUser("i", "Jane", "jane@x.com", "m@x.com").Describe());
        Assert.Equal("Jane (m@x.com)", new ResolvedUser("i", "Jane", null, "m@x.com").Describe());
        Assert.Equal("(no name) (i)", new ResolvedUser("i", null, null, null).Describe());
    }
}
