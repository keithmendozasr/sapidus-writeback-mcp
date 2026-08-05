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

    [Test]
    public async Task CreateFolderAsync_real_mode_creates_the_missing_segment_and_returns_it()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                // No existing child - ResolveFolderPathAsync's own resolve step and
                // CreateFolderPathAsync's existing-child lookup both see nothing here.
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"value":[]}""", Encoding.UTF8, "application/json"),
                });
            }

            Assert.That(request.Method, Is.EqualTo(HttpMethod.Post));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"new-folder-id","eTag":"\"etag-value\""}""", Encoding.UTF8, "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var service = CreateService(handler, options);

        var result = await service.CreateFolderAsync("new-folder", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.False);
            Assert.That(result.AlreadyExisted, Is.False);
            Assert.That(result.Item?.Id, Is.EqualTo("new-folder-id"));
        });
    }

    [Test]
    public async Task CreateFileAsync_real_mode_creates_the_file_and_logs_an_audit_entry()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Put));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"new-file-id","eTag":"\"etag-value\""}""", Encoding.UTF8, "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var logger = new RecordingLogger<DriveWriteService>();
        var service = CreateService(handler, options, logger);

        var result = await service.CreateFileAsync("notes.md", "content", conflictBehavior: "replace", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.False);
            Assert.That(result.Item?.Id, Is.EqualTo("new-file-id"));
            Assert.That(logger.Messages, Has.Some.Contains("new-file-id"), "the audit log entry should include the created item's id.");
        });
    }

    [Test]
    public async Task UpdateFileContentAsync_dry_run_resolves_the_target_and_makes_no_mutating_call()
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
        var service = CreateService(handler, DriveWriteOptions.Default);

        var result = await service.UpdateFileContentAsync("notes.md", "new content", "\"old-etag\"", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.True);
            Assert.That(result.Item?.Id, Is.EqualTo("existing-id"));
        });
    }

    [Test]
    public async Task UpdateFileContentAsync_real_mode_resolves_then_sends_an_If_Match_PUT()
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

            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Put));
                Assert.That(request.Headers.GetValues("If-Match").Single(), Is.EqualTo("\"old-etag\""));
            });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"existing-id","eTag":"\"new-etag\""}""", Encoding.UTF8, "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var service = CreateService(handler, options);

        var result = await service.UpdateFileContentAsync("notes.md", "new content", "\"old-etag\"", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.False);
            Assert.That(result.Item?.ETag, Is.EqualTo("\"new-etag\""));
        });
    }

    [Test]
    public async Task RenameItemAsync_dry_run_resolves_the_target_and_makes_no_mutating_call()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method != HttpMethod.Get)
                throw new InvalidOperationException($"Dry-run must never send a {request.Method} request.");

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"existing-id","name":"old-name.md"}""", Encoding.UTF8, "application/json"),
            });
        });
        var service = CreateService(handler, DriveWriteOptions.Default);

        var result = await service.RenameItemAsync("old-name.md", "new-name.md", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.True);
            Assert.That(result.NewName, Is.EqualTo("new-name.md"));
            Assert.That(result.Item?.Id, Is.EqualTo("existing-id"));
        });
    }

    [Test]
    public async Task RenameItemAsync_real_mode_resolves_then_sends_a_PATCH()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"id":"existing-id","name":"old-name.md"}""", Encoding.UTF8, "application/json"),
                });
            }

            Assert.That(request.Method, Is.EqualTo(HttpMethod.Patch));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"existing-id","name":"new-name.md"}""", Encoding.UTF8, "application/json"),
            });
        });
        var options = new DriveWriteOptions(DryRun: false, MaxContentBytes: 1_048_576);
        var service = CreateService(handler, options);

        var result = await service.RenameItemAsync("old-name.md", "new-name.md", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.False);
            Assert.That(result.Item?.Name, Is.EqualTo("new-name.md"));
        });
    }

    [Test]
    public async Task MoveItemAsync_dry_run_resolves_item_and_destination_and_makes_no_mutating_call()
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
        var service = CreateService(handler, DriveWriteOptions.Default);

        var result = await service.MoveItemAsync("notes.md", "target-folder", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.True);
            Assert.That(result.NewParentId, Is.EqualTo("parent-id"));
            Assert.That(result.Item?.Id, Is.EqualTo("item-id"));
        });
    }

    [Test]
    public async Task MoveItemAsync_real_mode_resolves_then_sends_a_PATCH()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            var uri = request.RequestUri!.ToString();

            if (request.Method == HttpMethod.Get && uri.Contains("target-folder"))
            {
                // Service-layer resolution of the destination parent by path.
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
                // DriveGraphClient.MoveItemAsync's own destination-by-id lookup for the
                // same-drive defense-in-depth check.
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
        var service = CreateService(handler, options);

        var result = await service.MoveItemAsync("notes.md", "target-folder", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result.DryRun, Is.False);
            Assert.That(result.NewParentId, Is.EqualTo("parent-id"));
            Assert.That(result.Item?.Id, Is.EqualTo("item-id"));
        });
    }
}
