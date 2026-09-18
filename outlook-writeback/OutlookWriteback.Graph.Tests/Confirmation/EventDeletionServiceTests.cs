using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Graph.Confirmation;
using OutlookWriteback.Graph.Tests.TestSupport;
using Sapidus.Writeback.Shared.Confirmation;

namespace OutlookWriteback.Graph.Tests.Confirmation;

/// <summary>
/// Exercises EventDeletionService's two-step, batch-shaped gating against a stubbed HTTP
/// handler - same fast, offline tier as OutlookGraphClientIntegrationTests. The point of these
/// tests is proving the DELETEs only ever fire behind a valid confirmation token, that one
/// failing event doesn't block the rest of the batch, and that a confirming call may name any
/// subset of the originally previewed IDs.
/// </summary>
[TestFixture]
[Category("Integration")]
public class EventDeletionServiceTests
{
    private const string NotFoundBody =
        """{"error":{"code":"ErrorItemNotFound","message":"The specified object was not found."}}""";

    private static EventDeletionService CreateService(StubHttpMessageHandler handler, DateTimeOffset now) =>
        CreateService(handler, new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now)));

    private static EventDeletionService CreateService(StubHttpMessageHandler handler, ConfirmationTokenService tokenService)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));

        return new EventDeletionService(graphClient, tokenService);
    }

    private static string LastPathSegment(HttpRequestMessage request) => request.RequestUri!.AbsolutePath.Split('/')[^1];

    [Test]
    public async Task RequestDeletionAsync_returns_details_for_every_previewed_event_and_one_shared_token()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Get));
            var eventId = LastPathSegment(request);

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent($$"""{"id":"{{eventId}}","subject":"Standup"}""", Encoding.UTF8, "application/json"),
            });
        });

        var preview = await CreateService(handler, DateTimeOffset.UtcNow).RequestDeletionAsync(["event-1", "event-2"]);

        Assert.Multiple(() =>
        {
            Assert.That(preview.Events, Has.Count.EqualTo(2));
            Assert.That(preview.Events, Has.All.Matches<PreviewedEventDeletion>(e => e.Found));
            Assert.That(preview.ConfirmationToken, Is.Not.Null.And.Not.Empty);
        });
    }

    [Test]
    public async Task RequestDeletionAsync_marks_an_event_not_found_without_failing_the_rest()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            var eventId = LastPathSegment(request);

            if (eventId == "missing-event")
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(NotFoundBody, Encoding.UTF8, "application/json"),
                });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent($$"""{"id":"{{eventId}}","subject":"Standup"}""", Encoding.UTF8, "application/json"),
            });
        });

        var preview = await CreateService(handler, DateTimeOffset.UtcNow).RequestDeletionAsync(["found-event", "missing-event"]);

        Assert.Multiple(() =>
        {
            Assert.That(preview.Events.Single(e => e.EventId == "found-event").Found, Is.True);
            Assert.That(preview.Events.Single(e => e.EventId == "missing-event").Found, Is.False);
        });
    }

    [Test]
    public void RequestDeletionAsync_throws_when_more_than_the_max_batch_size_is_requested()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException(
            "Graph must not be called for a request that exceeds the batch size cap."));

        var tooManyIds = Enumerable.Range(0, EventDeletionService.MaxBatchSize + 1).Select(i => $"event-{i}").ToArray();

        Assert.ThrowsAsync<ArgumentException>(() =>
            CreateService(handler, DateTimeOffset.UtcNow).RequestDeletionAsync(tooManyIds));
    }

    [Test]
    public async Task ConfirmDeletionAsync_deletes_all_previewed_events_when_the_full_set_is_reconfirmed()
    {
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(DateTimeOffset.UtcNow));
        var token = tokenService.IssueBatch(["event-1", "event-2"]);

        var deletedIds = new List<string>();
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Delete));
            deletedIds.Add(LastPathSegment(request));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent));
        });

        var result = await CreateService(handler, tokenService).ConfirmDeletionAsync(["event-1", "event-2"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.True);
            Assert.That(result.Outcomes, Has.All.Matches<EventDeletionOutcome>(o => o.Deleted));
            Assert.That(deletedIds, Is.EquivalentTo(new[] { "event-1", "event-2" }));
        });
    }

    [Test]
    public async Task ConfirmDeletionAsync_only_deletes_the_resent_subset_when_some_ids_are_dropped()
    {
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(DateTimeOffset.UtcNow));
        var token = tokenService.IssueBatch(["event-1", "event-2", "event-3"]);

        var deletedIds = new List<string>();
        var handler = new StubHttpMessageHandler(request =>
        {
            deletedIds.Add(LastPathSegment(request));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent));
        });

        var result = await CreateService(handler, tokenService).ConfirmDeletionAsync(["event-1", "event-3"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.True);
            Assert.That(result.Outcomes.Select(o => o.EventId), Is.EquivalentTo(new[] { "event-1", "event-3" }));
            Assert.That(deletedIds, Is.EquivalentTo(new[] { "event-1", "event-3" }));
        });
    }

    [Test]
    public async Task ConfirmDeletionAsync_rejects_an_id_outside_the_originally_previewed_set_without_failing_the_rest()
    {
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(DateTimeOffset.UtcNow));
        var token = tokenService.IssueBatch(["event-1"]);

        var deletedIds = new List<string>();
        var handler = new StubHttpMessageHandler(request =>
        {
            deletedIds.Add(LastPathSegment(request));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent));
        });

        var result = await CreateService(handler, tokenService).ConfirmDeletionAsync(["event-1", "not-previewed"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.True);
            Assert.That(result.Outcomes.Single(o => o.EventId == "event-1").Deleted, Is.True);
            Assert.That(result.Outcomes.Single(o => o.EventId == "not-previewed").Deleted, Is.False);
            Assert.That(deletedIds, Is.EqualTo(new[] { "event-1" }));
        });
    }

    [Test]
    public async Task ConfirmDeletionAsync_reports_a_failure_for_one_event_without_blocking_the_rest()
    {
        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(DateTimeOffset.UtcNow));
        var token = tokenService.IssueBatch(["event-1", "event-2"]);

        var handler = new StubHttpMessageHandler(request =>
        {
            if (LastPathSegment(request) == "event-1")
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent(NotFoundBody, Encoding.UTF8, "application/json"),
                });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent));
        });

        var result = await CreateService(handler, tokenService).ConfirmDeletionAsync(["event-1", "event-2"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.True);
            Assert.That(result.Outcomes.Single(o => o.EventId == "event-1").Deleted, Is.False);
            Assert.That(result.Outcomes.Single(o => o.EventId == "event-2").Deleted, Is.True);
        });
    }

    [Test]
    public async Task ConfirmDeletionAsync_does_not_call_Graph_when_the_token_is_invalid()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException(
            "Graph must not be called when the confirmation token fails validation."));

        var result = await CreateService(handler, DateTimeOffset.UtcNow)
            .ConfirmDeletionAsync(["event-1"], "not-a-valid-token");

        Assert.That(result.TokenValid, Is.False);
    }
}
