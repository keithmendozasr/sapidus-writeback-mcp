using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Functions;
using OutlookWriteback.Graph;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class UpdateEventToolTests
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
        var tool = new UpdateEventTool(CreateClient(handler));

        var result = await tool.RunAsync(null!, "AAkA-fake-event-id", null, null, null, null, null, null, ["alice@example.com", "   "], null, null, null, null, null, null, null, null);

        Assert.That(result, Does.StartWith("No changes were made.").And.Contain("must not be blank"));
    }

    [Test]
    public async Task RunAsync_clears_attendees_when_an_empty_array_is_provided()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            Assert.That(body, Does.Contain("\"attendees\":[]"));

            return new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
            };
        });
        var tool = new UpdateEventTool(CreateClient(handler));

        await tool.RunAsync(null!, "AAkA-fake-event-id", null, null, null, null, null, null, [], null, null, null, null, null, null, null, null);
    }

    private static (UpdateEventTool Tool, Func<string> LastBody) CreateRecordingTool()
    {
        var lastBody = string.Empty;
        var handler = new StubHttpMessageHandler(async request =>
        {
            lastBody = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            return new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"AAkA-fake-event-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        return (new UpdateEventTool(CreateClient(handler)), () => lastBody);
    }

    [Test]
    public async Task RunAsync_sends_a_recurrence_when_recurrenceType_is_provided_with_start_and_timeZone()
    {
        var (tool, lastBody) = CreateRecordingTool();

        // 2026-08-11 is the second Tuesday of August 2026.
        await tool.RunAsync(
            null!, "AAkA-fake-event-id", null, "2026-08-11T09:00:00-04:00", "2026-08-11T10:00:00-04:00", "America/New_York",
            null, null, null, null,
            "monthly", null, ["tuesday"], "second", null, null, null);

        Assert.Multiple(() =>
        {
            Assert.That(lastBody(), Does.Contain("\"type\":\"relativeMonthly\""));
            Assert.That(lastBody(), Does.Contain("\"index\":\"second\""));
            Assert.That(lastBody(), Does.Not.Contain("\"subject\""));
        });
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_recurrenceType_is_provided_without_start_and_timeZone()
    {
        var (tool, lastBody) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "AAkA-fake-event-id", null, null, null, null,
            null, null, null, null,
            "daily", null, null, null, null, null, null);

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("No changes were made.").And.Contain("start and timeZone"));
            Assert.That(lastBody(), Is.Empty);
        });
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_a_recurrence_parameter_is_given_without_recurrenceType()
    {
        var (tool, _) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "AAkA-fake-event-id", null, "2026-08-04T09:00:00-04:00", null, "America/New_York",
            null, null, null, null,
            null, 2, null, null, null, null, null);

        Assert.That(result, Does.StartWith("No changes were made.").And.Contain("recurrenceType is required"));
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_nothing_to_update_is_provided()
    {
        var (tool, _) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "AAkA-fake-event-id", null, null, null, null,
            null, null, null, null,
            null, null, null, null, null, null, null);

        Assert.That(result, Does.StartWith("No changes were made.").And.Contain("At least one of").And.Contain("reminderMinutes"));
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_start_is_not_a_timestamp()
    {
        var (tool, _) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "AAkA-fake-event-id", null, "next tuesday", null, "America/New_York",
            null, null, null, null,
            null, null, null, null, null, null, null);

        Assert.That(result, Does.StartWith("No changes were made.").And.Contain("start must be an ISO 8601"));
    }

    [Test]
    public async Task RunAsync_returns_a_correctable_message_when_timeZone_is_unknown()
    {
        var (tool, _) = CreateRecordingTool();

        var result = await tool.RunAsync(
            null!, "AAkA-fake-event-id", null, "2026-08-04T09:00:00-04:00", null, "Nowhere/Land",
            null, null, null, null,
            null, null, null, null, null, null, null);

        Assert.That(result, Does.StartWith("No changes were made.").And.Contain("isn't a recognized time zone"));
    }

    [Test]
    public async Task RunAsync_says_the_change_was_applied_to_the_series_when_given_an_occurrences_ID()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            var path = request.RequestUri!.AbsolutePath;
            var json = request.Method == HttpMethod.Patch
                ? """{"id":"AAkA-master-id"}"""
                : path.EndsWith("/AAkA-master-id")
                    ? """{"id":"AAkA-master-id","start":{"dateTime":"2026-10-09T17:00:00.0000000","timeZone":"America/Los_Angeles"}}"""
                    : """{"id":"AAkA-occurrence-id","type":"occurrence","seriesMasterId":"AAkA-master-id"}""";

            return Task.FromResult(new HttpResponseMessage(System.Net.HttpStatusCode.OK)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json"),
            });
        });
        var tool = new UpdateEventTool(CreateClient(handler));

        var result = await tool.RunAsync(
            null!, "AAkA-occurrence-id", null, "2026-10-09T17:00:00-07:00", null, "America/Los_Angeles",
            null, null, null, null,
            "daily", null, null, null, null, null, "2027-04-09");

        Assert.That(result, Does.StartWith("Event updated.").And.Contain("applied to the series itself").And.Contain("AAkA-master-id"));
    }

    [Test]
    public async Task RunAsync_accepts_the_values_as_the_MCP_extension_re_renders_them()
    {
        var (tool, lastBody) = CreateRecordingTool();

        // The exact failing call from the field: start/recurrenceUntil arrive re-rendered by the MCP extension.
        var result = await tool.RunAsync(
            null!, "AAkA-fake-event-id", null, "10/09/2026 17:00:00 -07:00", null, "America/Los_Angeles",
            null, null, null, null,
            "daily", null, null, null, null, null, "04/09/2027 00:00:00 -07:00");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.StartWith("Event updated."));
            Assert.That(lastBody(), Does.Contain("\"startDate\":\"2026-10-09\"").And.Contain("\"endDate\":\"2027-04-09\""));
        });
    }
}
