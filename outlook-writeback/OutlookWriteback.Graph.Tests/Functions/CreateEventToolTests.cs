using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Functions;
using OutlookWriteback.Graph;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class CreateEventToolTests
{
    private static OutlookGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new OutlookGraphClient(graphClient);
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_an_attendee_entry_is_blank()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph should not be called."));
        var tool = new CreateEventTool(CreateClient(handler));

        var result = await tool.RunAsync(
                null!,
                "Standup",
                "2026-08-01T09:00:00-05:00",
                "2026-08-01T09:30:00-05:00",
                "UTC",
                null,
                null,
                ["alice@example.com", "   "],
                null,
                null, null, null, null, null, null, null);

        Assert.That(result, Does.StartWith("No event was created.").And.Contain("must not be blank"));
    }

    [Test]
    public async Task RunAsync_trims_whitespace_before_sending_to_graph()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            Assert.That(body, Does.Contain("\"address\":\"alice@example.com\""));

            return new HttpResponseMessage(System.Net.HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
            };
        });
        var tool = new CreateEventTool(CreateClient(handler));

        await tool.RunAsync(
            null!,
            "Standup",
            "2026-08-01T09:00:00-05:00",
            "2026-08-01T09:30:00-05:00",
            "UTC",
            null,
            null,
            ["  alice@example.com  "],
            null,
            null, null, null, null, null, null, null);
    }

    [Test]
    public async Task RunAsync_creates_an_event_with_no_attendees_when_attendees_is_omitted()
    {
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(System.Net.HttpStatusCode.Created)
        {
            Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
        }));
        var tool = new CreateEventTool(CreateClient(handler));

        var result = await tool.RunAsync(
            null!,
            "Standup",
            "2026-08-01T09:00:00-05:00",
            "2026-08-01T09:30:00-05:00",
            "UTC",
            null,
            null,
            null,
            null,
            null, null, null, null, null, null, null);

        Assert.That(result, Does.Contain("AAkA-fake-event-id"));
        Assert.That(result, Does.Not.Contain("series"));
    }

    private static (CreateEventTool Tool, Func<string> LastBody) CreateRecordingTool()
    {
        var lastBody = string.Empty;
        var handler = new StubHttpMessageHandler(async request =>
        {
            lastBody = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            return new HttpResponseMessage(System.Net.HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        return (new CreateEventTool(CreateClient(handler)), () => lastBody);
    }

    [Test]
    public async Task RunAsync_sends_a_recurrence_when_recurrenceType_is_provided()
    {
        var (tool, lastBody) = CreateRecordingTool();

        // 2026-08-04 is a Tuesday.
        var result = await tool.RunAsync(
            null!, "PT", "2026-08-04T09:00:00-04:00", "2026-08-04T10:00:00-04:00", "America/New_York",
            null, null, null, null,
            "weekly", 2, ["tuesday", "thursday"], null, null, 10, null);

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("Recurring event series created").And.Contain("AAkA-fake-event-id"));
            Assert.That(lastBody(), Does.Contain("\"recurrence\""));
            Assert.That(lastBody(), Does.Contain("\"type\":\"weekly\""));
            Assert.That(lastBody(), Does.Contain("\"numberOfOccurrences\":10"));
        });
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_a_recurrence_parameter_is_given_without_recurrenceType()
    {
        var (tool, lastBody) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "PT", "2026-08-04T09:00:00-04:00", "2026-08-04T10:00:00-04:00", "America/New_York",
            null, null, null, null,
            null, 2, null, null, null, null, null);

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("No event was created.").And.Contain("recurrenceType is required"));
            Assert.That(lastBody(), Is.Empty, "Graph must not be called for rejected input.");
        });
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_the_recurrence_is_invalid()
    {
        var (tool, _) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "PT", "2026-08-04T09:00:00-04:00", "2026-08-04T10:00:00-04:00", "America/New_York",
            null, null, null, null,
            "weekly", null, null, null, null, null, null);

        Assert.That(result, Does.StartWith("No event was created.").And.Contain("recurrenceDaysOfWeek is required"));
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_recurrenceUntil_is_not_yyyy_MM_dd()
    {
        var (tool, lastBody) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "PT", "2026-08-04T09:00:00-04:00", "2026-08-04T10:00:00-04:00", "America/New_York",
            null, null, null, null,
            "daily", null, null, null, null, null, "04/09/2027 00:00:00 -07:00");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("No event was created.").And.Contain("yyyy-MM-dd").And.Contain("2027-04-09"));
            Assert.That(lastBody(), Is.Empty);
        });
    }

    [TestCase("tomorrow at 9", "start must be an ISO 8601")]
    [TestCase("", "start must be an ISO 8601")]
    [TestCase("10am", "start must be an ISO 8601")]
    [TestCase("2026-08-04T09:00:00", "start must be an ISO 8601")]
    [TestCase("08/04/2026 09:00:00 -04:00", "start must be an ISO 8601")]
    public async Task RunAsync_returns_a_correctable_message_when_start_is_not_a_timestamp(string start, string expected)
    {
        var (tool, lastBody) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "PT", start, "2026-08-04T10:00:00-04:00", "America/New_York",
            null, null, null, null,
            null, null, null, null, null, null, null);

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("No event was created.").And.Contain(expected).And.Contain("2026-08-01T09:00:00-05:00"));
            Assert.That(lastBody(), Is.Empty);
        });
    }

    [TestCase("2026-08-04T09:00-04:00")]
    [TestCase("2026-08-04T09:00:00-04:00")]
    [TestCase("2026-08-04T13:00:00Z")]
    [TestCase("2026-08-04T09:00:00.123-04:00")]
    public async Task RunAsync_accepts_ISO_8601_timestamps_with_an_offset(string start)
    {
        var (tool, lastBody) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "PT", start, "2026-08-04T14:00:00Z", "America/New_York",
            null, null, null, null,
            null, null, null, null, null, null, null);

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("Event created."));
            Assert.That(lastBody(), Does.Contain("2026-08-04T09:00:00"));
        });
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_end_is_not_a_timestamp()
    {
        var (tool, _) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "PT", "2026-08-04T09:00:00-04:00", "10am", "America/New_York",
            null, null, null, null,
            null, null, null, null, null, null, null);

        Assert.That(result, Does.Contain("end must be an ISO 8601"));
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_timeZone_is_unknown()
    {
        var (tool, lastBody) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "PT", "2026-08-04T09:00:00-04:00", "2026-08-04T10:00:00-04:00", "Eastern",
            null, null, null, null,
            null, null, null, null, null, null, null);

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("No event was created.").And.Contain("'Eastern' isn't a recognized time zone").And.Contain("America/New_York"));
            Assert.That(lastBody(), Is.Empty);
        });
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_reminderMinutes_is_negative()
    {
        var (tool, _) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "PT", "2026-08-04T09:00:00-04:00", "2026-08-04T10:00:00-04:00", "America/New_York",
            null, null, null, -5,
            null, null, null, null, null, null, null);

        Assert.That(result, Does.Contain("reminderMinutes must not be negative"));
    }
}
