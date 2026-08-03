using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using DriveWriteback.Graph.Tests.TestSupport;

namespace DriveWriteback.Graph.Tests;

/// <summary>
/// Fast, offline tier: exercises DriveGraphClient's real request-building and response
/// deserialization through the Graph SDK, against a stubbed HTTP handler instead of the
/// network. No credentials, no live tenant, safe for CI. For real checks against the
/// actual Graph API, see DriveGraphClientE2ETests.
/// </summary>
[TestFixture]
[Category("Integration")]
public class DriveGraphClientIntegrationTests
{
    private static DriveGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new DriveGraphClient(graphClient);
    }
}
