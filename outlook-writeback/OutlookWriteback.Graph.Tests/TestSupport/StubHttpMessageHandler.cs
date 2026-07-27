namespace OutlookWriteback.Graph.Tests.TestSupport;

/// <summary>
/// Fakes the HTTP boundary Microsoft.Graph's Kiota-generated client sends requests through, so
/// tests exercise real request-building and response deserialization without a network call.
/// </summary>
internal sealed class StubHttpMessageHandler : HttpMessageHandler
{
    private readonly Func<HttpRequestMessage, HttpResponseMessage> _respond;

    public StubHttpMessageHandler(Func<HttpRequestMessage, HttpResponseMessage> respond)
    {
        _respond = respond;
    }

    public HttpRequestMessage? LastRequest { get; private set; }

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        LastRequest = request;

        return Task.FromResult(_respond(request));
    }
}
