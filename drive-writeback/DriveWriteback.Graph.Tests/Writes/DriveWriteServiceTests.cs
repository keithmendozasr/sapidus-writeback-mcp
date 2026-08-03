using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using DriveWriteback.Graph.Tests.TestSupport;
using DriveWriteback.Graph.Writes;

namespace DriveWriteback.Graph.Tests.Writes;

[TestFixture]
[Category("Unit")]
public class DriveWriteServiceTests
{
    private static DriveWriteService CreateService(StubHttpMessageHandler handler, DriveWriteOptions options, RecordingLogger<DriveWriteService>? logger = null)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());
        var client = new DriveGraphClient(graphClient);

        return new DriveWriteService(client, options, logger ?? new RecordingLogger<DriveWriteService>());
    }
}
