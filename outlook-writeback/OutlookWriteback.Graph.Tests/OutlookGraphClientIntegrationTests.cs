using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests;

/// <summary>
/// Fast, offline tier: exercises OutlookGraphClient's real request-building and response
/// deserialization through the Graph SDK, against a stubbed HTTP handler instead of the
/// network. No credentials, no live tenant, safe for CI. For real checks against the actual
/// Graph API, see OutlookGraphClientE2ETests.
/// </summary>
[TestFixture]
[Category("Integration")]
public class OutlookGraphClientIntegrationTests
{
    private static OutlookGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new OutlookGraphClient(graphClient);
    }
}
