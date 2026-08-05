using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using DriveWriteback.Graph.Confirmation;
using DriveWriteback.Graph.Tests.TestSupport;
using DriveWriteback.Graph.Writes;
using Sapidus.Writeback.Shared.Confirmation;

namespace DriveWriteback.Graph.Tests.Confirmation;

/// <summary>
/// Exercises DriveItemDeletionService's two-step gating against a stubbed HTTP handler - same
/// fast, offline tier as DriveGraphClientIntegrationTests. The point of these tests is proving
/// the DELETE only ever fires behind a valid token, matching expected_name, and a folder that's
/// still empty (or recursive=true).
/// </summary>
[TestFixture]
[Category("Integration")]
public class DriveItemDeletionServiceTests
{
    private static (DriveItemDeletionService Service, ConfirmationTokenService TokenService) CreateService(
        StubHttpMessageHandler handler,
        DateTimeOffset now,
        bool dryRun = false,
        ConfirmationTokenService? tokenService = null)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());
        var client = new DriveGraphClient(graphClient);
        var options = new DriveWriteOptions(DryRun: dryRun, MaxContentBytes: 1_048_576);
        var writeService = new DriveWriteService(client, options, new RecordingLogger<DriveWriteService>());
        var resolvedTokenService = tokenService
            ?? new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));

        return (new DriveItemDeletionService(client, writeService, resolvedTokenService), resolvedTokenService);
    }

    [Test]
    public async Task RequestDeletionAsync_returns_the_items_details_and_a_token_without_deleting()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Get));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"item-id","name":"notes.md"}""", Encoding.UTF8, "application/json"),
            });
        });

        var (service, _) = CreateService(handler, DateTimeOffset.UtcNow);
        var pending = await service.RequestDeletionAsync("notes.md", "notes.md", recursive: false, driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(pending.ItemId, Is.EqualTo("item-id"));
            Assert.That(pending.Name, Is.EqualTo("notes.md"));
            Assert.That(pending.ConfirmationToken, Is.Not.Null.And.Not.Empty);
        });
    }

    [Test]
    public void RequestDeletionAsync_throws_DriveItemNameMismatchException_when_expected_name_does_not_match()
    {
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent("""{"id":"item-id","name":"notes.md"}""", Encoding.UTF8, "application/json"),
        }));

        var (service, _) = CreateService(handler, DateTimeOffset.UtcNow);

        Assert.That(
            () => service.RequestDeletionAsync("notes.md", "a-different-name.md", recursive: false, driveId: "drive-id"),
            Throws.InstanceOf<DriveItemNameMismatchException>());
    }

    [Test]
    public void RequestDeletionAsync_throws_DriveFolderNotEmptyException_for_a_non_empty_folder_without_recursive()
    {
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(
                """{"id":"folder-id","name":"reports","folder":{"childCount":3}}""",
                Encoding.UTF8,
                "application/json"),
        }));

        var (service, _) = CreateService(handler, DateTimeOffset.UtcNow);

        Assert.That(
            () => service.RequestDeletionAsync("reports", "reports", recursive: false, driveId: "drive-id"),
            Throws.InstanceOf<DriveFolderNotEmptyException>());
    }

    [Test]
    public async Task ConfirmDeletionAsync_deletes_when_the_token_is_valid_for_the_same_item()
    {
        var now = DateTimeOffset.UtcNow;
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));
        var token = tokenService.Issue("item-id");

        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                // ConfirmDeletionAsync's own re-fetch, used to re-check expected_name/recursive.
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"id":"item-id","name":"notes.md"}""", Encoding.UTF8, "application/json"),
                });
            }

            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Delete));
                Assert.That(request.RequestUri!.ToString(), Does.Contain("items/item-id"));
            });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent));
        });

        var (service, _) = CreateService(handler, now, tokenService: tokenService);

        var confirmed = await service.ConfirmDeletionAsync("item-id", token, "notes.md", recursive: false, driveId: "drive-id");

        Assert.That(confirmed, Is.EqualTo(ConfirmedItemDeletion.Deleted));
    }

    [Test]
    public async Task ConfirmDeletionAsync_does_not_call_Graph_when_the_token_is_invalid()
    {
        var handler = new StubHttpMessageHandler(
            _ => throw new InvalidOperationException("Graph must not be called when the confirmation token fails validation."));

        var (service, _) = CreateService(handler, DateTimeOffset.UtcNow);

        var confirmed = await service.ConfirmDeletionAsync("item-id", "not-a-valid-token", "notes.md", recursive: false, driveId: "drive-id");

        Assert.That(confirmed, Is.EqualTo(ConfirmedItemDeletion.InvalidToken));
    }

    [Test]
    public async Task ConfirmDeletionAsync_returns_Deleted_when_the_item_is_already_gone()
    {
        var now = DateTimeOffset.UtcNow;
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));
        var token = tokenService.Issue("item-id");

        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
        {
            Content = new StringContent(
                """{"error":{"code":"itemNotFound","message":"Item not found"}}""",
                Encoding.UTF8,
                "application/json"),
        }));

        var (service, _) = CreateService(handler, now, tokenService: tokenService);

        // PRD §7 idempotency: confirming a delete for an item already gone must not error.
        var confirmed = await service.ConfirmDeletionAsync("item-id", token, "notes.md", recursive: false, driveId: "drive-id");

        Assert.That(confirmed, Is.EqualTo(ConfirmedItemDeletion.Deleted));
    }

    [Test]
    public async Task ConfirmDeletionAsync_returns_DryRun_and_makes_no_DELETE_call_when_dry_run_is_on()
    {
        var now = DateTimeOffset.UtcNow;
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));
        var token = tokenService.Issue("item-id");

        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method != HttpMethod.Get)
                throw new InvalidOperationException($"Dry-run must never send a {request.Method} request.");

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"item-id","name":"notes.md"}""", Encoding.UTF8, "application/json"),
            });
        });

        var (service, _) = CreateService(handler, now, dryRun: true, tokenService: tokenService);

        var confirmed = await service.ConfirmDeletionAsync("item-id", token, "notes.md", recursive: false, driveId: "drive-id");

        Assert.That(confirmed, Is.EqualTo(ConfirmedItemDeletion.DryRun));
    }
}
