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

    [Test]
    public async Task RunAsync_real_mode_reports_the_move_and_the_item_id()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            var uri = request.RequestUri!.ToString();

            if (request.Method == HttpMethod.Get && uri.Contains("target-folder"))
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"id":"parent-id","name":"target-folder","folder":{}}""", Encoding.UTF8, "application/json"),
                });
            }

            if (request.Method == HttpMethod.Get && uri.Contains("notes.md"))
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"id":"item-id","name":"notes.md"}""", Encoding.UTF8, "application/json"),
                });
            }

            if (request.Method == HttpMethod.Get)
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"id":"parent-id","parentReference":{"driveId":"drive-id"}}""", Encoding.UTF8, "application/json"),
                });
            }

            Assert.That(request.Method, Is.EqualTo(HttpMethod.Patch));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"item-id","name":"notes.md"}""", Encoding.UTF8, "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var tool = CreateTool(handler, options);

        var result = await tool.RunAsync(null!, "notes.md", "target-folder", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("moved into \"target-folder\""));
            Assert.That(result, Does.Contain("parent-id"));
            Assert.That(result, Does.Contain("item-id"));
        });
    }
}
