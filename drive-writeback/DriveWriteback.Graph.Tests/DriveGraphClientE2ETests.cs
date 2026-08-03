namespace DriveWriteback.Graph.Tests;

/// <summary>
/// True end-to-end tier: real Graph API, real OneDrive. Requires the "Drive Writeback MCP"
/// Entra app (Files.ReadWrite.All + Sites.Read.All, admin consent granted - see
/// ../DEPLOYMENT.md) and is skipped via Assert.Ignore when the environment variables below
/// aren't set - it cannot run in CI or without the user's own tenant credentials.
///
/// Each test in this fixture answers one open Phase 0 validation-spike question from
/// docs/active/PRD-drive-write.md §11 and cleans up after itself so repeated runs don't
/// accumulate junk in the real OneDrive.
/// </summary>
[TestFixture]
[Category("E2E")]
public class DriveGraphClientE2ETests
{
    private DriveGraphClient? _client;

    [SetUp]
    public void SetUp()
    {
        var tenantId = Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_TENANT_ID");
        var clientId = Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_CLIENT_ID");

        if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId))
            Assert.Ignore("Set DRIVE_WRITEBACK_TENANT_ID and DRIVE_WRITEBACK_CLIENT_ID (the 'Drive Writeback MCP' Entra app) to run Phase 0 live Graph checks.");

        _client = DriveGraphClient.CreateWithInteractiveBrowserAuth(
            tenantId!,
            clientId!,
            ["Files.ReadWrite.All", "Sites.Read.All"]);
    }
}
