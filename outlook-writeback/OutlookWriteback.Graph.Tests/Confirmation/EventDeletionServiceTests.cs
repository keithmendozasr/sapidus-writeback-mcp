using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Graph.Confirmation;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests.Confirmation;

/// <summary>
/// Exercises EventDeletionService's two-step gating against a stubbed HTTP handler - same fast,
/// offline tier as OutlookGraphClientIntegrationTests. The point of these tests is proving the
/// DELETE only ever fires behind a valid confirmation token, per PRD.md section 6.
/// </summary>
[TestFixture]
[Category("Integration")]
public class EventDeletionServiceTests
{
    private static EventDeletionService CreateService(StubHttpMessageHandler handler, DateTimeOffset now)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));
        var tokenService = new DeleteConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));

        return new EventDeletionService(graphClient, tokenService);
    }

    [Test]
    public async Task RequestDeletionAsync_returns_the_events_details_and_a_token_without_deleting()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Get));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"id":"AAkA-fake-event-id","subject":"Standup"}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });

        var pending = await CreateService(handler, DateTimeOffset.UtcNow).RequestDeletionAsync("AAkA-fake-event-id");

        Assert.Multiple(() =>
        {
            Assert.That(pending.EventId, Is.EqualTo("AAkA-fake-event-id"));
            Assert.That(pending.Subject, Is.EqualTo("Standup"));
            Assert.That(pending.ConfirmationToken, Is.Not.Null.And.Not.Empty);
        });
    }
}
