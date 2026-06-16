using System.Net;
using System.Text;

namespace M365Migrate.Core.Tests;

/// <summary>
/// A test <see cref="HttpMessageHandler"/> that responds via a caller-supplied
/// function, capturing each request and its body for assertions.
/// </summary>
public sealed class FakeHttpMessageHandler : HttpMessageHandler
{
    private readonly Func<HttpRequestMessage, string, HttpResponseMessage> _responder;

    public List<HttpRequestMessage> Requests { get; } = new();
    public List<string> Bodies { get; } = new();

    public FakeHttpMessageHandler(Func<HttpRequestMessage, string, HttpResponseMessage> responder)
        => _responder = responder;

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var body = request.Content is null ? "" : await request.Content.ReadAsStringAsync(cancellationToken);
        Requests.Add(request);
        Bodies.Add(body);
        return _responder(request, body);
    }

    public static HttpResponseMessage Json(HttpStatusCode code, string json) =>
        new(code) { Content = new StringContent(json, Encoding.UTF8, "application/json") };

    public static HttpResponseMessage Text(HttpStatusCode code, string text) =>
        new(code) { Content = new StringContent(text, Encoding.UTF8, "text/plain") };
}
