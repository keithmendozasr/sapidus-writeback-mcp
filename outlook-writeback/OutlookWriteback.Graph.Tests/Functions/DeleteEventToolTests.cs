using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Functions;
using OutlookWriteback.Graph.Confirmation;
using OutlookWriteback.Graph.Tests.TestSupport;
using Sapidus.Writeback.Shared.Confirmation;

namespace OutlookWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class DeleteEventToolTests
{
    private static DeleteEventTool CreateTool(
        StubHttpMessageHandler handler,
        ConfirmationTokenService? tokenService = null,
        RecordingLogger<DeleteEventTool>? logger = null)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));
        var deletionService = new EventDeletionService(graphClient, tokenService ?? new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key")));

        return new DeleteEventTool(deletionService, logger ?? new RecordingLogger<DeleteEventTool>());
    }

    [Test]
    public async Task RunAsync_requires_at_least_one_event_id()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph must not be called."));
        var tool = CreateTool(handler);

        var result = await tool.RunAsync(null!, [], null);

        Assert.That(result, Is.EqualTo("At least one eventId is required."));
    }

    [Test]
    public async Task RunAsync_rejects_more_than_the_max_batch_size_without_calling_the_service()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph must not be called."));
        var tool = CreateTool(handler);
        var tooManyIds = Enumerable.Range(0, EventDeletionService.MaxBatchSize + 1).Select(i => $"event-{i}").ToArray();

        var result = await tool.RunAsync(null!, tooManyIds, null);

        Assert.That(result, Does.Contain(EventDeletionService.MaxBatchSize.ToString()));
    }

    [Test]
    public async Task RunAsync_reports_when_the_confirmation_token_is_invalid()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph must not be called."));
        var tool = CreateTool(handler);

        var result = await tool.RunAsync(null!, ["event-1"], "not-a-valid-token");

        Assert.That(result, Does.Contain("invalid or expired"));
    }

    [Test]
    public async Task RunAsync_previews_found_and_not_found_events_without_deleting()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            var eventId = request.RequestUri!.AbsolutePath.Split('/')[^1];

            if (eventId == "missing-event")
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(
                        """{"error":{"code":"ErrorItemNotFound","message":"not found"}}""",
                        Encoding.UTF8,
                        "application/json"),
                });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent($$"""{"id":"{{eventId}}","subject":"Standup"}""", Encoding.UTF8, "application/json"),
            });
        });
        var tool = CreateTool(handler);

        var result = await tool.RunAsync(null!, ["found-event", "missing-event"], null);

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("Standup"));
            Assert.That(result, Does.Contain("NOT FOUND: missing-event"));
            Assert.That(result, Does.Contain("This has NOT been deleted yet"));
        });
    }

    [Test]
    public async Task RunAsync_logs_preview_request_count_with_event_ids_kept_out_of_the_message()
    {
        var handler = new StubHttpMessageHandler(request => Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
        {
            Content = new StringContent("""{"error":{"code":"ErrorItemNotFound","message":"not found"}}""", Encoding.UTF8, "application/json"),
        }));
        var logger = new RecordingLogger<DeleteEventTool>();
        var tool = CreateTool(handler, logger: logger);

        await tool.RunAsync(null!, ["event-1", "event-2"], null);

        Assert.Multiple(() =>
        {
            Assert.That(logger.Messages, Has.Some.Contains("requested-count=2"));
            Assert.That(logger.Messages, Has.None.Contains("event-1").Or.Contains("event-2"));

            var eventIds = logger.Properties.SelectMany(p => p).Where(kvp => kvp.Key == "EventIds").Select(kvp => kvp.Value);

            Assert.That(eventIds, Has.Some.EqualTo(new[] { "event-1", "event-2" }));
        });
    }

    [Test]
    public async Task RunAsync_confirms_and_summarizes_deleted_and_failed_events()
    {
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"));
        var token = tokenService.IssueBatch(["event-1", "event-2"]);

        var handler = new StubHttpMessageHandler(request =>
        {
            var eventId = request.RequestUri!.AbsolutePath.Split('/')[^1];

            if (eventId == "event-1")
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(
                        """{"error":{"code":"ErrorItemNotFound","message":"already deleted"}}""",
                        Encoding.UTF8,
                        "application/json"),
                });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent));
        });
        var tool = CreateTool(handler, tokenService);

        var result = await tool.RunAsync(null!, ["event-1", "event-2"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("1 of 2 deleted."));
            Assert.That(result, Does.Contain("event-1"));
        });
    }

    [Test]
    public async Task RunAsync_logs_confirm_request_count_with_event_ids_kept_out_of_the_message()
    {
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"));
        var token = tokenService.IssueBatch(["event-1", "event-2"]);
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent)));
        var logger = new RecordingLogger<DeleteEventTool>();
        var tool = CreateTool(handler, tokenService, logger);

        await tool.RunAsync(null!, ["event-1", "event-2"], token);

        Assert.Multiple(() =>
        {
            Assert.That(logger.Messages, Has.Some.Contains("requested-count=2"));
            Assert.That(logger.Messages, Has.None.Contains("event-1").Or.Contains("event-2"));

            var eventIds = logger.Properties.SelectMany(p => p).Where(kvp => kvp.Key == "EventIds").Select(kvp => kvp.Value);

            Assert.That(eventIds, Has.Some.EqualTo(new[] { "event-1", "event-2" }));
        });
    }
}
