using System.Net;
using System.Text;
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

    [Test]
    public async Task CreateFolderAsync_dry_run_reports_no_missing_segments_and_makes_no_mutating_call()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method != HttpMethod.Get)
                throw new InvalidOperationException($"Dry-run must never send a {request.Method} request.");

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"value":[{"id":"folder-id","name":"existing","folder":{}}]}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });
        var service = CreateService(handler, DriveWriteOptions.Default);

        var result = await service.CreateFolderAsync("existing", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.True);
            Assert.That(result.AlreadyExisted, Is.True);
            Assert.That(result.CreatedSegments, Is.Empty);
        });
    }
}
