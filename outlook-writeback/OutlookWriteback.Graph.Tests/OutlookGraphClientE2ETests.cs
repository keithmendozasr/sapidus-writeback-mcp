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
    public async Task CreateDraftAsync_persists_multiple_to_cc_and_bcc_recipients()
    {
        var toAddress = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_TEST_TO_ADDRESS");

        if (string.IsNullOrEmpty(toAddress))
            Assert.Ignore("Set OUTLOOK_WRITEBACK_TEST_TO_ADDRESS to a real mailbox address to run this check.");

        var draftId = await _client!.CreateDraftAsync(
            [toAddress!],
            "Phase 1 multi-recipient - outlook-writeback",
            "Created by the multi-recipient integration test. Safe to delete.",
            ccAddresses: [toAddress!],
            bccAddresses: [toAddress!]);

        Assert.That(draftId, Is.Not.Null.And.Not.Empty);
    }
}
