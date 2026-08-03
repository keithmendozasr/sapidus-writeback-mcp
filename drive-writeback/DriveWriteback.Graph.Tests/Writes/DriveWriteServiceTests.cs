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

    [Test]
    public async Task CreateFileAsync_dry_run_reports_target_already_exists_without_writing()
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
        var service = CreateService(handler, DriveWriteOptions.Default);

        var result = await service.CreateFileAsync("notes.md", "content", conflictBehavior: "rename", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.True);
            Assert.That(result.TargetAlreadyExisted, Is.True);
            Assert.That(result.Item, Is.Null);
        });
    }

    [Test]
    public void CreateFileAsync_dry_run_throws_DriveParentNotFoundException_when_the_parent_is_missing()
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
        var service = CreateService(handler, DriveWriteOptions.Default);

        Assert.That(
            () => service.CreateFileAsync("sub/notes.md", "content", conflictBehavior: "replace", driveId: "drive-id"),
            Throws.InstanceOf<DriveParentNotFoundException>());
    }

    [Test]
    public void CreateFileAsync_throws_DriveContentTooLargeException_before_any_network_call()
    {
        var handler = new StubHttpMessageHandler(
            _ => throw new InvalidOperationException("Content-size validation must fail before any Graph call - dry-run or not."));
        var options = new DriveWriteOptions(DryRun: true, MaxContentBytes: 5);
        var service = CreateService(handler, options);

        Assert.That(
            () => service.CreateFileAsync("notes.md", "this content is definitely over five bytes", driveId: "drive-id"),
            Throws.InstanceOf<DriveContentTooLargeException>());
    }
}
