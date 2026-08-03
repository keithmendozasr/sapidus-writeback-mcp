using DriveWriteback.Graph.Auth;
using DriveWriteback.Graph.Tests.TestSupport;

namespace DriveWriteback.Graph.Tests.Auth;

[TestFixture]
[Category("Unit")]
public class GraphTokenEndpointClientTests
{
    private static GraphTokenEndpointClient CreateClient(StubHttpMessageHandler handler) =>
        new("tenant-id", "client-id", new HttpClient(handler));
}
