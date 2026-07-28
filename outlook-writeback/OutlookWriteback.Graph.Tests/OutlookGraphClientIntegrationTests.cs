using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests;

/// <summary>
/// Fast, offline tier: exercises OutlookGraphClient's real request-building and response
/// deserialization through the Graph SDK, against a stubbed HTTP handler instead of the
/// network. No credentials, no live tenant, safe for CI. For real checks against the actual
/// Graph API, see OutlookGraphClientE2ETests.
/// </summary>
[TestFixture]
[Category("Integration")]
public class OutlookGraphClientIntegrationTests
{
    private static OutlookGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new OutlookGraphClient(graphClient);
    }

    [Test]
    public async Task CreateDraftAsync_sends_a_POST_to_me_messages_and_returns_the_created_id()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Post));
                Assert.That(request.RequestUri!.AbsolutePath, Does.Contain("/me/messages"));
            });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAMk-fake-draft-id"}""", Encoding.UTF8, "application/json"),
            });
        });

        var draftId = await CreateClient(handler).CreateDraftAsync("owner@example.com", "Subject", "Body");

        Assert.That(draftId, Is.EqualTo("AAMk-fake-draft-id"));
    }

    [Test]
    public async Task CreateDraftAsync_serializes_html_content_type_when_isHtml_is_true()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            Assert.That(body, Does.Contain("\"contentType\":\"html\""));

            return new HttpResponseMessage(HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAMk-fake-draft-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        await CreateClient(handler).CreateDraftAsync(
            "owner@example.com",
            "Subject",
            "<table></table>",
            isHtml: true);
    }

    [Test]
    public async Task UpdateDraftAsync_sends_a_PATCH_to_me_messages_id_and_returns_the_updated_id()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Patch));
                Assert.That(request.RequestUri!.AbsolutePath, Does.Contain("/me/messages/AAMk-fake-draft-id"));
                Assert.That(body, Does.Not.Contain("attachments"));
            });

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"AAMk-fake-draft-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        var updatedId = await CreateClient(handler).UpdateDraftAsync("AAMk-fake-draft-id", subject: "New subject");

        Assert.That(updatedId, Is.EqualTo("AAMk-fake-draft-id"));
    }

    [Test]
    public async Task CreateEventAsync_sends_a_POST_to_me_events_and_returns_the_created_id()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Post));
                Assert.That(request.RequestUri!.AbsolutePath, Does.Contain("/me/events"));
            });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
            });
        });

        var start = DateTimeOffset.UtcNow.AddDays(1);
        var eventId = await CreateClient(handler).CreateEventAsync("Standup", start, start.AddHours(1), "UTC");

        Assert.That(eventId, Is.EqualTo("AAkA-fake-event-id"));
    }

    [Test]
    public async Task CreateEventAsync_serializes_attendees_when_provided()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            Assert.That(body, Does.Contain("alice@example.com"));

            return new HttpResponseMessage(HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        var start = DateTimeOffset.UtcNow.AddDays(1);

        await CreateClient(handler).CreateEventAsync(
            "Standup",
            start,
            start.AddHours(1),
            "UTC",
            attendeeAddresses: ["alice@example.com"]);
    }

    [Test]
    public async Task CreateEventAsync_serializes_a_non_utc_timeZone_with_no_offset_suffix()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            Assert.Multiple(() =>
            {
                Assert.That(body, Does.Contain("\"timeZone\":\"America/New_York\""));
                Assert.That(body, Does.Not.Match("\"dateTime\":\"[^\"]*Z\""));
            });

            return new HttpResponseMessage(HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        var start = DateTimeOffset.UtcNow.AddDays(1);

        await CreateClient(handler).CreateEventAsync(
            "Standup",
            start,
            start.AddHours(1),
            "America/New_York");
    }

    [Test]
    public async Task UpdateEventAsync_sends_a_PATCH_to_me_events_id_and_returns_the_updated_id()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Patch));
                Assert.That(request.RequestUri!.AbsolutePath, Does.Contain("/me/events/AAkA-fake-event-id"));
                Assert.That(body, Does.Contain("New subject"));
            });

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        var updatedId = await CreateClient(handler).UpdateEventAsync("AAkA-fake-event-id", subject: "New subject");

        Assert.That(updatedId, Is.EqualTo("AAkA-fake-event-id"));
    }

    [Test]
    public async Task DeleteEventAsync_sends_a_DELETE_to_me_events_id()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Delete));
                Assert.That(request.RequestUri!.AbsolutePath, Does.Contain("/me/events/AAkA-fake-event-id"));
            });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NoContent));
        });

        await CreateClient(handler).DeleteEventAsync("AAkA-fake-event-id");
    }

    [Test]
    public async Task GetEventByIdAsync_sends_a_GET_to_me_events_id_and_returns_the_resolved_event()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Get));
                Assert.That(request.RequestUri!.AbsolutePath, Does.Contain("/me/events/AAkA-connector-event-id"));
            });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"id":"AAkA-connector-event-id","subject":"Standup"}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });

        var resolved = await CreateClient(handler).GetEventByIdAsync("AAkA-connector-event-id");

        Assert.That(resolved?.Id, Is.EqualTo("AAkA-connector-event-id"));
    }
}
