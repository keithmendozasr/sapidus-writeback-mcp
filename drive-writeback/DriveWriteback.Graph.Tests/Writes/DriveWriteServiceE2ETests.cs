using Microsoft.Extensions.Logging.Abstractions;
using DriveWriteback.Graph.Writes;

namespace DriveWriteback.Graph.Tests.Writes;

/// <summary>
/// True end-to-end tier for the Phase 1 write surface: real Graph API, real OneDrive.
/// Requires the "Drive Writeback MCP" Entra app (see ../DEPLOYMENT.md) and self-skips via
/// Assert.Ignore when the environment variables below aren't set. Complements
/// DriveGraphClientE2ETests, which exercises DriveGraphClient directly - this fixture goes
/// through DriveWriteService, the layer create_folder/create_file's tool classes actually
/// call.
///
/// As of this commit, none of these tests have been run against a live tenant (no browser
/// session available in this implementation pass) - see docs/active/PRD-drive-write.md §11
/// for where findings get written back once someone does run them.
/// </summary>
[TestFixture]
[Category("E2E")]
public class DriveWriteServiceE2ETests
{
    private DriveGraphClient? _client;

    [SetUp]
    public void SetUp()
    {
        var tenantId = Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_TENANT_ID");
        var clientId = Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_CLIENT_ID");

        if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId))
            Assert.Ignore("Set DRIVE_WRITEBACK_TENANT_ID and DRIVE_WRITEBACK_CLIENT_ID (the 'Drive Writeback MCP' Entra app) to run Phase 1 live DriveWriteService checks.");

        _client = DriveGraphClient.CreateWithInteractiveBrowserAuth(
            tenantId!,
            clientId!,
            ["Files.ReadWrite.All", "Sites.Read.All"]);
    }

    private DriveWriteService CreateService(bool dryRun) =>
        new(_client!, new DriveWriteOptions(dryRun, MaxContentBytes: 1_048_576), NullLogger<DriveWriteService>.Instance);
}
