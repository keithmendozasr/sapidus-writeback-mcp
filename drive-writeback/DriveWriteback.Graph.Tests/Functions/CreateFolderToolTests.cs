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
public class CreateFolderToolTests
{
    private static CreateFolderTool CreateTool(StubHttpMessageHandler handler, DriveWriteOptions options)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());
        var client = new DriveGraphClient(graphClient);
        var writeService = new DriveWriteService(client, options, new RecordingLogger<DriveWriteService>());

        return new CreateFolderTool(writeService);
    }

    [Test]
    public async Task RunAsync_dry_run_reports_the_would_be_created_segment_and_makes_no_mutating_call()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method != HttpMethod.Get)
                throw new InvalidOperationException($"Dry-run must never send a {request.Method} request.");

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"value":[]}""", Encoding.UTF8, "application/json"),
            });
        });
        var tool = CreateTool(handler, DriveWriteOptions.Default);

        var result = await tool.RunAsync(null!, "new-folder", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("[DRY RUN]"));
            Assert.That(result, Does.Contain("new-folder"));
            Assert.That(result, Does.Contain("Not created"));
        });
    }

    [Test]
    public async Task RunAsync_dry_run_reports_already_exists_when_nothing_is_missing()
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
        var tool = CreateTool(handler, DriveWriteOptions.Default);

        var result = await tool.RunAsync(null!, "existing", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("[DRY RUN]"));
            Assert.That(result, Does.Contain("already exists"));
        });
    }

    [Test]
    public async Task RunAsync_real_mode_reports_the_ready_folder_id_and_etag()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"value":[]}""", Encoding.UTF8, "application/json"),
                });
            }

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"new-folder-id","eTag":"\"etag-value\""}""", Encoding.UTF8, "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var tool = CreateTool(handler, options);

        var result = await tool.RunAsync(null!, "new-folder", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("ready"));
            Assert.That(result, Does.Contain("new-folder-id"));
            Assert.That(result, Does.Contain("etag-value"));
        });
    }

    [Test]
    public async Task RunAsync_real_mode_says_already_existed_instead_of_created_when_nothing_changed()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        """{"value":[{"id":"existing-folder-id","name":"existing","folder":{},"eTag":"\"etag-value\""}]}""",
                        Encoding.UTF8,
                        "application/json"),
                });
            }

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Conflict)
            {
                Content = new StringContent(
                    """{"error":{"code":"nameAlreadyExists","message":"An item with that name already exists."}}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var tool = CreateTool(handler, options);

        var result = await tool.RunAsync(null!, "existing", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("already existed"));
            Assert.That(result, Does.Not.Contain("(created)"));
            Assert.That(result, Does.Contain("existing-folder-id"));
        });
    }
}
