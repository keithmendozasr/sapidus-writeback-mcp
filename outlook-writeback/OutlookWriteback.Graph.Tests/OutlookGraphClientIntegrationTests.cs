using OutlookWriteback.Graph;

namespace OutlookWriteback.Graph.Tests;

/// <summary>
/// PRD Phase 0 live checks against the real Graph API. Require the "Outlook Writeback MCP"
/// Entra app (Mail.ReadWrite + Calendars.ReadWrite, admin consent granted) and are skipped
/// via Assert.Ignore when the environment variables below aren't set - they cannot run in CI
/// or without the user's own tenant credentials.
/// </summary>
[TestFixture]
[Category("Integration")]
public class OutlookGraphClientIntegrationTests
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
}
