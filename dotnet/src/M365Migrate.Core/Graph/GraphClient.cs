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

    private static HttpContent JsonBody(object body)
    {
        var json = JsonSerializer.Serialize(body);
        return new StringContent(json, Encoding.UTF8, "application/json");
    }
}
