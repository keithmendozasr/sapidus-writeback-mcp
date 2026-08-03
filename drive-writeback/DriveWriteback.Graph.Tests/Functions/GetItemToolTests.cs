using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using DriveWriteback.Functions;
using DriveWriteback.Graph;
using DriveWriteback.Graph.Tests.TestSupport;

namespace DriveWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class GetItemToolTests
{
    private static DriveGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new DriveGraphClient(graphClient);
    }
}
