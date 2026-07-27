using OutlookWriteback.Graph;

namespace OutlookWriteback.Graph.Tests;

/// <summary>
/// True end-to-end tier: real Graph API, real mailbox. Require the "Outlook Writeback MCP"
/// Entra app (Mail.ReadWrite + Calendars.ReadWrite, admin consent granted) and are skipped
/// via Assert.Ignore when the environment variables below aren't set - they cannot run in CI
/// or without the user's own tenant credentials. For fast, offline, network-free checks against
/// OutlookGraphClient's request/response handling, see OutlookGraphClientIntegrationTests.
/// </summary>
[TestFixture]
[Category("E2E")]
public class OutlookGraphClientE2ETests
{
    private OutlookGraphClient? _client;

    [SetUp]
    public void SetUp()
    {
        var tenantId = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_TENANT_ID");
        var clientId = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_CLIENT_ID");

        if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId))
            Assert.Ignore("Set OUTLOOK_WRITEBACK_TENANT_ID and OUTLOOK_WRITEBACK_CLIENT_ID (the 'Outlook Writeback MCP' Entra app) to run Phase 0 live Graph checks.");

        _client = OutlookGraphClient.CreateWithInteractiveBrowserAuth(
            tenantId!,
            clientId!,
            ["Mail.ReadWrite", "Calendars.ReadWrite"]);
    }

    [Test]
    public async Task CreateDraftAsync_persists_a_draft_message_and_returns_its_id()
    {
        var toAddress = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_TEST_TO_ADDRESS");

        if (string.IsNullOrEmpty(toAddress))
            Assert.Ignore("Set OUTLOOK_WRITEBACK_TEST_TO_ADDRESS to a real mailbox address to run this check.");

        var draftId = await _client!.CreateDraftAsync(
            toAddress!,
            "Phase 0 spike - outlook-writeback",
            "Created by the Phase 0 spike integration test. Safe to delete.");

        Assert.That(draftId, Is.Not.Null.And.Not.Empty);
    }

    [Test]
    public async Task CreateEventAsync_creates_an_event_and_returns_its_id()
    {
        var start = DateTimeOffset.UtcNow.AddDays(30);
        var end = start.AddHours(1);

        var eventId = await _client!.CreateEventAsync(
            "Phase 0 spike - outlook-writeback",
            start,
            end,
            bodyText: "Created by the Phase 0 spike integration test. Safe to delete.");

        Assert.That(eventId, Is.Not.Null.And.Not.Empty);
    }

    /// <summary>
    /// PRD open question #2: an event ID found by the M365 connector's calendar search must
    /// resolve through this app's own Graph credentials against the same /me/events ID space.
    /// </summary>
    [Test]
    public async Task GetEventByIdAsync_resolves_an_event_id_returned_by_the_M365_connector()
    {
        var connectorEventId = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_CONNECTOR_EVENT_ID");

        if (string.IsNullOrEmpty(connectorEventId))
            Assert.Ignore("Set OUTLOOK_WRITEBACK_CONNECTOR_EVENT_ID to a real event ID from the M365 connector's calendar search.");

        var resolved = await _client!.GetEventByIdAsync(connectorEventId!);

        Assert.That(resolved?.Id, Is.EqualTo(connectorEventId));
    }
}
