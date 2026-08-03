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
public class CreateFileToolTests
{
    private static CreateFileTool CreateTool(StubHttpMessageHandler handler, DriveWriteOptions options)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());
        var client = new DriveGraphClient(graphClient);
        var writeService = new DriveWriteService(client, options, new RecordingLogger<DriveWriteService>());

        return new CreateFileTool(writeService);
    }

    [Test]
    public async Task RunAsync_dry_run_reports_parent_confirmed_would_create()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method != HttpMethod.Get)
                throw new InvalidOperationException($"Dry-run must never send a {request.Method} request.");

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
            {
                Content = new StringContent(
                    """{"error":{"code":"itemNotFound","message":"Item not found"}}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });
        var tool = CreateTool(handler, DriveWriteOptions.Default);

        var result = await tool.RunAsync(null!, "notes.md", "content", conflictBehavior: null, driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("[DRY RUN]"));
            Assert.That(result, Does.Contain("Parent confirmed to exist"));
            Assert.That(result, Does.Contain("Not written"));
        });
    }

    [Test]
    public async Task RunAsync_dry_run_reports_target_already_exists_with_the_conflict_behavior()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method != HttpMethod.Get)
                throw new InvalidOperationException($"Dry-run must never send a {request.Method} request.");

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"existing-id"}""", Encoding.UTF8, "application/json"),
            });
        });
        var tool = CreateTool(handler, DriveWriteOptions.Default);

        var result = await tool.RunAsync(null!, "notes.md", "content", conflictBehavior: "rename", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("[DRY RUN]"));
            Assert.That(result, Does.Contain("already exists"));
            Assert.That(result, Does.Contain("conflict_behavior=rename"));
            Assert.That(result, Does.Contain("Not written"));
        });
    }

    [Test]
    public async Task RunAsync_real_mode_reports_id_size_etag_and_web_url()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Put));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"id":"new-file-id","name":"notes.md","size":7,"eTag":"\"etag-value\"","webUrl":"https://example/notes.md"}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var tool = CreateTool(handler, options);

        var result = await tool.RunAsync(null!, "notes.md", "content", conflictBehavior: "replace", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("created"));
            Assert.That(result, Does.Contain("new-file-id"));
            Assert.That(result, Does.Contain("7 bytes"));
            Assert.That(result, Does.Contain("etag-value"));
            Assert.That(result, Does.Contain("https://example/notes.md"));
        });
    }
}
