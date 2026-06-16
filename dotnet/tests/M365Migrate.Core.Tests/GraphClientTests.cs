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
    public async Task UploadLargeFile_ChunksWithContentRanges()
    {
        var ranges = new List<string>();
        var puts = 0;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (req.Method == HttpMethod.Post && uri.Contains("createUploadSession"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"uploadUrl\":\"https://upload.example/sess\"}");
            if (req.Method == HttpMethod.Put && uri.Contains("upload.example"))
            {
                if (req.Content!.Headers.TryGetValues("Content-Range", out var v))
                    ranges.Add(v.First());
                puts++;
                return puts == 1
                    ? FakeHttpMessageHandler.Json(HttpStatusCode.Accepted, "{\"nextExpectedRanges\":[\"4-\"]}")
                    : FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"done\"}");
            }
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{}");
        });

        var client = Build(handler);
        var result = await client.UploadLargeFileAsync(
            "/drive/root:/big.bin:/createUploadSession",
            System.Text.Encoding.ASCII.GetBytes("ABCDEF"),
            chunkSize: 4);

        Assert.Equal("done", result.GetStringOrNull("id"));
        Assert.Equal(new[] { "bytes 0-3/6", "bytes 4-5/6" }, ranges);
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
