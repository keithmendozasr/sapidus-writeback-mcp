namespace DriveWriteback.Graph.Tests.TestSupport;

/// <summary>
/// Fakes the HTTP boundary Microsoft.Graph's Kiota-generated client sends requests through, so
/// tests exercise real request-building and response deserialization without a network call.
/// </summary>
internal sealed class StubHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> respond) : HttpMessageHandler
{
    public HttpRequestMessage? LastRequest { get; private set; }

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        LastRequest = request;

        return respond(request);
    }
}
