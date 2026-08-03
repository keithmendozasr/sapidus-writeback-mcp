using DriveWriteback.Graph.Auth;
using DriveWriteback.Graph.Tests.TestSupport;

namespace DriveWriteback.Graph.Tests.Auth;

[TestFixture]
[Category("Unit")]
public class SilentGraphCredentialTests
{
    private static SilentGraphCredential CreateCredential(StubHttpMessageHandler handler, FakeRefreshTokenStore store)
    {
        var tokenEndpointClient = new GraphTokenEndpointClient("tenant-id", "client-id", new HttpClient(handler));

        return new SilentGraphCredential(tokenEndpointClient, store, ["Files.ReadWrite.All"]);
    }
}
