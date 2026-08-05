using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using DriveWriteback.Functions;
using DriveWriteback.Graph.Tests.TestSupport;
using DriveWriteback.Graph.Writes;

namespace DriveWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class MoveItemToolTests
{
    private static MoveItemTool CreateTool(StubHttpMessageHandler handler, DriveWriteOptions options)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());
        var client = new DriveGraphClient(graphClient);
        var writeService = new DriveWriteService(client, options, new RecordingLogger<DriveWriteService>());

        return new MoveItemTool(writeService);
    }

    [Test]
    public async Task RunAsync_dry_run_reports_would_move_without_moving()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method != HttpMethod.Get)
                throw new InvalidOperationException($"Dry-run must never send a {request.Method} request.");

            var uri = request.RequestUri!.ToString();
            var body = uri.Contains("target-folder")
                ? """{"id":"parent-id","name":"target-folder","folder":{}}"""
                : """{"id":"item-id","name":"notes.md"}""";

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json"),
            });
        });
        var tool = CreateTool(handler, DriveWriteOptions.Default);

        var result = await tool.RunAsync(null!, "notes.md", "target-folder", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("[DRY RUN]"));
            Assert.That(result, Does.Contain("item-id"));
            Assert.That(result, Does.Contain("parent-id"));
            Assert.That(result, Does.Contain("Not moved"));
        });
    }
}
