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
        bool dryRun = false)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());
        var client = new DriveGraphClient(graphClient);
        var options = new DriveWriteOptions(DryRun: dryRun, MaxContentBytes: 1_048_576);
        var writeService = new DriveWriteService(client, options, new RecordingLogger<DriveWriteService>());
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));

        return (new DriveItemDeletionService(client, writeService, tokenService), tokenService);
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
}
