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

    /// <summary>
    /// PRD §11 Phase 0: "Confirm mkdir -p approach... against a OneDrive... library."
    /// Creating the same nested path twice must succeed both times - the second run
    /// exercises the 409-as-already-exists path CreateFolderPathAsync depends on.
    /// </summary>
    [Test]
    public async Task CreateFolderPathAsync_is_idempotent_mkdir_p()
    {
        var spikeRoot = $"phase0-spike-mkdirp-{Guid.NewGuid():N}";
        var nestedPath = $"{spikeRoot}/a/b/c";

        try
        {
            var firstCreate = await _client!.CreateFolderPathAsync(nestedPath);
            Assert.That(firstCreate, Is.Not.Null);

            var secondCreate = await _client.CreateFolderPathAsync(nestedPath);
            Assert.That(secondCreate, Is.Not.Null, "Re-creating an existing path should succeed, not throw.");

            var resolved = await _client.GetItemByPathAsync(nestedPath);
            Assert.That(resolved?.Folder, Is.Not.Null, "Expected the deepest segment to be a folder.");
        }
        finally
        {
            await _client!.DeleteItemAsync(spikeRoot);
        }
    }
}
