using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using M365Migrate.Core.Auth;

namespace M365Migrate.Core.Graph;

/// <summary>
/// A thin Microsoft Graph REST client: bearer-token injection, paging over
/// <c>@odata.nextLink</c>, and retry with backoff on throttling (429) and
/// transient 5xx responses.
///
/// The <see cref="HttpClient"/> and <see cref="TokenProvider"/> are injected so
/// the client can be unit-tested with a fake handler and no network.
/// </summary>
public sealed class GraphClient
{
    public const string DefaultBaseUrl = "https://graph.microsoft.com/v1.0";
    private static readonly HashSet<int> Retryable = new() { 429, 500, 502, 503, 504 };

    // Upload-session chunks must be a multiple of 320 KiB (except the final chunk).
    public const int UploadFragment = 320 * 1024;
    public const int DefaultUploadChunk = 10 * UploadFragment; // 3.2 MiB per PUT

    private readonly HttpClient _http;
    private readonly TokenProvider _tokenProvider;
    private readonly int _maxRetries;
    private readonly Func<int, Task> _delay;

    public GraphClient(
        HttpClient http,
        TokenProvider tokenProvider,
        string baseUrl = DefaultBaseUrl,
        int maxRetries = 5,
        Func<int, Task>? delay = null)
    {
        _http = http;
        _tokenProvider = tokenProvider;
        BaseUrl = baseUrl.TrimEnd('/');
        _maxRetries = maxRetries;
        _delay = delay ?? (ms => Task.Delay(ms));
    }

    /// <summary>The Graph base URL this client targets (e.g. for building $ref links).</summary>
    public string BaseUrl { get; }

    private string Absolute(string url) =>
        url.StartsWith("http://") || url.StartsWith("https://")
            ? url
            : $"{BaseUrl}/{url.TrimStart('/')}";

    /// <summary>Issue a single request, retrying on throttling/transient errors.</summary>
    public async Task<HttpResponseMessage> SendAsync(
        HttpMethod method,
        string url,
        Func<HttpContent>? contentFactory = null,
        CancellationToken ct = default)
    {
        var target = Absolute(url);
        var attempt = 0;
        while (true)
        {
            using var request = new HttpRequestMessage(method, target);
            var token = await _tokenProvider(ct);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            if (contentFactory is not null)
                request.Content = contentFactory();

            var response = await _http.SendAsync(request, ct);
            var code = (int)response.StatusCode;

            if (Retryable.Contains(code) && attempt < _maxRetries)
            {
                var wait = RetryAfterMs(response, attempt);
                response.Dispose();
                await _delay(wait);
                attempt++;
                continue;
            }

            if (code >= 400)
            {
                var body = await response.Content.ReadAsStringAsync(ct);
                response.Dispose();
                throw new GraphException(code, body);
            }

            return response;
        }
    }

    private static int RetryAfterMs(HttpResponseMessage response, int attempt)
    {
        if (response.Headers.TryGetValues("Retry-After", out var values))
        {
            var raw = values.FirstOrDefault();
            if (double.TryParse(raw, out var seconds))
                return (int)(seconds * 1000);
        }
        // Exponential backoff capped at 60s.
        return (int)(Math.Min(Math.Pow(2, attempt), 60) * 1000);
    }

    /// <summary>GET a single resource and return its parsed JSON root element.</summary>
    public async Task<JsonElement> GetAsync(string url, CancellationToken ct = default)
    {
        using var response = await SendAsync(HttpMethod.Get, url, null, ct);
        var text = await response.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(text);
        return doc.RootElement.Clone();
    }

    /// <summary>Yield every item across all pages of a Graph collection.</summary>
    public async Task<List<JsonElement>> GetAllAsync(string url, CancellationToken ct = default)
    {
        var items = new List<JsonElement>();
        string? next = url;
        while (next is not null)
        {
            using var response = await SendAsync(HttpMethod.Get, next, null, ct);
            var text = await response.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(text);
            var root = doc.RootElement;
            if (root.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
                foreach (var element in value.EnumerateArray())
                    items.Add(element.Clone());
            next = root.TryGetProperty("@odata.nextLink", out var link) ? link.GetString() : null;
        }
        return items;
    }

    /// <summary>
    /// Create a message from its raw MIME (RFC 5322) content. Graph expects the MIME
    /// base64-encoded with Content-Type text/plain. Returns the created message JSON.
    /// </summary>
    public async Task<JsonElement> PostMimeMessageAsync(string url, byte[] mime, CancellationToken ct = default)
    {
        var base64 = Convert.ToBase64String(mime);
        using var response = await SendAsync(HttpMethod.Post, url, () =>
        {
            var content = new StringContent(base64);
            content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");
            return content;
        }, ct);
        var text = await response.Content.ReadAsStringAsync(ct);
        if (string.IsNullOrWhiteSpace(text))
            return default;
        using var doc = JsonDocument.Parse(text);
        return doc.RootElement.Clone();
    }

    /// <summary>
    /// POST a JSON body and return the response's Location (or Content-Location)
    /// header — used for async creates (e.g. a migration-mode team) that return 202.
    /// </summary>
    public async Task<string?> PostJsonGetLocationAsync(string url, object body, CancellationToken ct = default)
    {
        using var response = await SendAsync(HttpMethod.Post, url, () => JsonBody(body), ct);
        if (response.Headers.Location is not null)
            return response.Headers.Location.ToString();
        if (response.Headers.TryGetValues("Location", out var values))
            return values.FirstOrDefault();
        return response.Content.Headers.ContentLocation?.ToString();
    }

    /// <summary>POST with no request body (e.g. teams completeMigration).</summary>
    public async Task PostNoBodyAsync(string url, CancellationToken ct = default)
    {
        using var response = await SendAsync(HttpMethod.Post, url, null, ct);
        _ = response;
    }

    /// <summary>POST a JSON body and return the parsed response (or a JSON null if empty).</summary>
    public async Task<JsonElement> PostJsonAsync(string url, object body, CancellationToken ct = default)
    {
        using var response = await SendAsync(HttpMethod.Post, url, () => JsonBody(body), ct);
        var text = await response.Content.ReadAsStringAsync(ct);
        if (string.IsNullOrWhiteSpace(text))
            return default;
        using var doc = JsonDocument.Parse(text);
        return doc.RootElement.Clone();
    }

    /// <summary>PATCH a JSON body (Graph returns 204 No Content on success).</summary>
    public async Task PatchJsonAsync(string url, object body, CancellationToken ct = default)
    {
        using var response = await SendAsync(HttpMethod.Patch, url, () => JsonBody(body), ct);
        _ = response;
    }

    /// <summary>PUT a JSON body (used for reference links like manager/$ref).</summary>
    public async Task PutJsonAsync(string url, object body, CancellationToken ct = default)
    {
        using var response = await SendAsync(HttpMethod.Put, url, () => JsonBody(body), ct);
        _ = response;
    }

    /// <summary>GET a resource and return its raw bytes (e.g. driveItem content).</summary>
    public async Task<byte[]> GetBytesAsync(string url, CancellationToken ct = default)
    {
        using var response = await SendAsync(HttpMethod.Get, url, null, ct);
        return await response.Content.ReadAsByteArrayAsync(ct);
    }

    /// <summary>PUT raw bytes (small-file upload) and return the parsed response.</summary>
    public async Task<JsonElement> PutBytesAsync(string url, byte[] data, CancellationToken ct = default)
    {
        using var response = await SendAsync(HttpMethod.Put, url, () =>
        {
            var content = new ByteArrayContent(data);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            return content;
        }, ct);
        var text = await response.Content.ReadAsStringAsync(ct);
        if (string.IsNullOrWhiteSpace(text))
            return default;
        using var doc = JsonDocument.Parse(text);
        return doc.RootElement.Clone();
    }

    /// <summary>
    /// Upload <paramref name="data"/> via a resumable upload session, in chunks.
    /// The upload URL Graph returns is pre-authenticated, so the bearer token is
    /// deliberately not sent on the chunk PUTs.
    /// </summary>
    public async Task<JsonElement> UploadLargeFileAsync(
        string createSessionUrl,
        byte[] data,
        string conflictBehavior = "replace",
        int chunkSize = DefaultUploadChunk,
        CancellationToken ct = default)
    {
        var session = await PostJsonAsync(
            createSessionUrl,
            new Dictionary<string, object>
            {
                ["item"] = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = conflictBehavior,
                },
            },
            ct);
        var uploadUrl = session.GetStringOrNull("uploadUrl");
        if (uploadUrl is null)
            throw new GraphException(500, "createUploadSession returned no uploadUrl");

        var total = data.Length;
        var start = 0;
        string lastBody = "";
        while (start < total)
        {
            var end = Math.Min(start + chunkSize, total);
            var length = end - start;
            var chunk = new byte[length];
            Array.Copy(data, start, chunk, 0, length);

            var content = new ByteArrayContent(chunk);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            content.Headers.TryAddWithoutValidation("Content-Range", $"bytes {start}-{end - 1}/{total}");

            using var response = await _http.PutAsync(uploadUrl, content, ct);
            var code = (int)response.StatusCode;
            lastBody = await response.Content.ReadAsStringAsync(ct);
            if (code >= 400)
                throw new GraphException(code, lastBody);
            start = end;
        }

        if (!string.IsNullOrWhiteSpace(lastBody))
        {
            using var doc = JsonDocument.Parse(lastBody);
            return doc.RootElement.Clone();
        }
        return default;
    }

    private static HttpContent JsonBody(object body)
    {
        var json = JsonSerializer.Serialize(body);
        return new StringContent(json, Encoding.UTF8, "application/json");
    }
}
