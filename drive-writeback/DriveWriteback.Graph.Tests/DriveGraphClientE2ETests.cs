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

    /// <summary>
    /// PRD §11/§12 Q5 Phase 0: "Confirm eTag vs cTag semantics for If-Match on /content
    /// - these differ, and picking wrong yields either false 412s or no protection at all."
    /// Uses two independent fresh files (one per tag) so a successful PUT against one
    /// doesn't invalidate the still-untested tag on the other. Also confirms a
    /// deliberately wrong tag reliably 412s, since a false-positive success there would
    /// mean update_file_content's mandatory if_match isn't actually protecting anything.
    /// </summary>
    [Test]
    public async Task ReplaceTextContentAsync_reports_which_tag_If_Match_honors()
    {
        var wrongTagRejected = await TryReplaceFreshItemAsync(tagSelector: null);
        var eTagHonored = await TryReplaceFreshItemAsync(tagSelector: item => item.ETag);
        var cTagHonored = await TryReplaceFreshItemAsync(tagSelector: item => item.CTag);

        TestContext.Out.WriteLine(
            $"PRD §12 Q5 finding: eTag If-Match {(eTagHonored ? "SUCCEEDED" : "412'd")}; "
            + $"cTag If-Match {(cTagHonored ? "SUCCEEDED" : "412'd")}.");

        Assert.That(wrongTagRejected, Is.False, "A deliberately wrong If-Match value must 412, not succeed.");
        Assert.That(
            eTagHonored || cTagHonored,
            Is.True,
            "At least one of eTag/cTag must be honored by If-Match on /content, or update_file_content's concurrency story needs rethinking.");
    }

    /// <summary>
    /// Creates a fresh file, attempts a content replacement using either a real tag
    /// (via tagSelector) or a deliberately wrong one (tagSelector: null), and reports
    /// whether the PUT succeeded. Always cleans up the file it created.
    /// </summary>
    private async Task<bool> TryReplaceFreshItemAsync(Func<Microsoft.Graph.Models.DriveItem, string?>? tagSelector)
    {
        var path = $"phase0-spike-tag-{Guid.NewGuid():N}.txt";

        try
        {
            await _client!.UploadTextContentAsync(path, "v1");

            string tag;

            if (tagSelector is null)
            {
                tag = "\"this-is-not-a-real-tag\"";
            }
            else
            {
                var initial = await _client.GetItemByPathAsync(path);
                var selected = tagSelector(initial!);
                Assert.That(selected, Is.Not.Null.And.Not.Empty);
                tag = selected!;
            }

            try
            {
                await _client.ReplaceTextContentAsync(path, "v2", tag);
                return true;
            }
            catch (DriveItemConcurrencyException)
            {
                return false;
            }
        }
        finally
        {
            await _client!.DeleteItemAsync(path);
        }
    }
}
