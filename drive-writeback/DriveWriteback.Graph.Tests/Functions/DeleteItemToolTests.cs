using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using DriveWriteback.Functions;
using DriveWriteback.Graph.Confirmation;
using DriveWriteback.Graph.Tests.TestSupport;
using DriveWriteback.Graph.Writes;
using Sapidus.Writeback.Shared.Confirmation;

namespace DriveWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class DeleteItemToolTests
{
    private static (DeleteItemTool Tool, ConfirmationTokenService TokenService) CreateTool(
        StubHttpMessageHandler handler,
        DateTimeOffset now,
        ConfirmationTokenService? tokenService = null)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());
        var client = new DriveGraphClient(graphClient);
        var writeService = new DriveWriteService(client, DriveWriteOptions.Default, new RecordingLogger<DriveWriteService>());
        var resolvedTokenService = tokenService
            ?? new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));
        var deletionService = new DriveItemDeletionService(client, writeService, resolvedTokenService);

        return (new DeleteItemTool(deletionService), resolvedTokenService);
    }

    [Test]
    public async Task RunAsync_without_a_token_previews_and_does_not_delete()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Get));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"item-id","name":"notes.md"}""", Encoding.UTF8, "application/json"),
            });
        });
        var (tool, _) = CreateTool(handler, DateTimeOffset.UtcNow);

        var result = await tool.RunAsync(null!, "notes.md", "notes.md", recursive: null, confirmationToken: null, driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("About to delete"));
            Assert.That(result, Does.Contain("notes.md"));
            Assert.That(result, Does.Contain("item-id"));
            Assert.That(result, Does.Contain("NOT been deleted yet"));
        });
    }

    [Test]
    public async Task RunAsync_with_a_valid_token_deletes()
    {
        var now = DateTimeOffset.UtcNow;
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));
        var token = tokenService.Issue("item-id");

        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"id":"item-id","name":"notes.md"}""", Encoding.UTF8, "application/json"),
                });
            }

            Assert.That(request.Method, Is.EqualTo(HttpMethod.Delete));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent));
        });
        var (tool, _) = CreateTool(handler, now, tokenService);

        var result = await tool.RunAsync(null!, "item-id", "notes.md", recursive: null, confirmationToken: token, driveId: "drive-id");

        Assert.That(result, Does.Contain("deleted"));
    }

    [Test]
    public async Task RunAsync_with_an_invalid_token_reports_invalid_or_expired()
    {
        var handler = new StubHttpMessageHandler(
            _ => throw new InvalidOperationException("Graph must not be called when the confirmation token fails validation."));
        var (tool, _) = CreateTool(handler, DateTimeOffset.UtcNow);

        var result = await tool.RunAsync(null!, "item-id", "notes.md", recursive: null, confirmationToken: "not-a-valid-token", driveId: "drive-id");

        Assert.That(result, Does.Contain("invalid or expired"));
    }
}
