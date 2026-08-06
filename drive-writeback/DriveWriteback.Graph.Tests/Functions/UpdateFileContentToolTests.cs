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
public class UpdateFileContentToolTests
{
    private static UpdateFileContentTool CreateTool(StubHttpMessageHandler handler, DriveWriteOptions options)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());
        var client = new DriveGraphClient(graphClient);
        var writeService = new DriveWriteService(client, options, new RecordingLogger<DriveWriteService>());

        return new UpdateFileContentTool(writeService);
    }

    [Test]
    public async Task RunAsync_dry_run_reports_would_update_without_writing()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method != HttpMethod.Get)
                throw new InvalidOperationException($"Dry-run must never send a {request.Method} request.");

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"existing-id","eTag":"\"old-etag\""}""", Encoding.UTF8, "application/json"),
            });
        });
        var tool = CreateTool(handler, DriveWriteOptions.Default);

        var result = await tool.RunAsync(null!, "notes.md", "new content", "\"old-etag\"", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("[DRY RUN]"));
            Assert.That(result, Does.Contain("existing-id"));
            Assert.That(result, Does.Contain("Not written"));
        });
    }

    [Test]
    public async Task RunAsync_real_mode_reports_id_size_and_new_etag()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"id":"existing-id","eTag":"\"old-etag\""}""", Encoding.UTF8, "application/json"),
                });
            }

            Assert.That(request.Headers.GetValues("If-Match").Single(), Is.EqualTo("\"old-etag\""));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"id":"existing-id","size":11,"eTag":"\"new-etag\""}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var tool = CreateTool(handler, options);

        var result = await tool.RunAsync(null!, "notes.md", "new content", "\"old-etag\"", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("updated"));
            Assert.That(result, Does.Contain("existing-id"));
            Assert.That(result, Does.Contain("11 bytes"));
            Assert.That(result, Does.Contain("new-etag"));
        });
    }
}
