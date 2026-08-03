using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Graph;
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

    /// <summary>
    /// PRD §7 "Idempotency"/§4.4: create_file's three conflict_behavior values against a
    /// real, already-existing file - fail must throw, replace must overwrite in place
    /// (same item id), rename must produce a second, non-colliding item.
    /// </summary>
    [Test]
    public async Task CreateFileAsync_real_mode_honors_all_three_conflict_behavior_values()
    {
        var service = CreateService(dryRun: false);
        var path = $"phase1-spike-conflict-{Guid.NewGuid():N}.txt";
        var renamedName = (string?)null;

        try
        {
            var original = await service.CreateFileAsync(path, "v1", conflictBehavior: "fail");
            Assert.That(original.Item?.Id, Is.Not.Null.And.Not.Empty);

            Assert.That(
                () => service.CreateFileAsync(path, "v2", conflictBehavior: "fail"),
                Throws.InstanceOf<DriveItemAlreadyExistsException>(),
                "fail must throw when the target already exists.");

            var replaced = await service.CreateFileAsync(path, "v3", conflictBehavior: "replace");
            Assert.That(replaced.Item?.Id, Is.EqualTo(original.Item!.Id), "replace must overwrite the same item, not create a new one.");

            var renamed = await service.CreateFileAsync(path, "v4", conflictBehavior: "rename");
            Assert.That(renamed.Item?.Id, Is.Not.EqualTo(original.Item.Id), "rename must produce a second, distinct item.");
            renamedName = renamed.Item?.Name;
        }
        finally
        {
            await _client!.DeleteItemAsync(path);

            // Root-level path, so the item's own Name is the deletable path.
            if (renamedName is not null)
                await _client.DeleteItemAsync(renamedName);
        }
    }

    /// <summary>
    /// Diagnostic, not a gate: does a content PUT to a path whose parent is missing
    /// auto-vivify the parent, or 404? DriveGraphClient.CreateFileAsync's own
    /// parent-existence guard makes this moot for the tool's own behavior (it never
    /// reaches Graph in that case, per DriveParentNotFoundException) - this bypasses that
    /// guard via RawClient to observe Graph's actual behavior directly, worth recording as
    /// a fact regardless, same as the Phase 0 colon-path bug was (PRD §11).
    /// </summary>
    [Test]
    public async Task ContentPut_to_a_missing_parent_is_an_open_question()
    {
        var raw = _client!.RawClient;
        var driveId = (await raw.Me.Drive.GetAsync())!.Id!;
        var missingParent = $"phase1-spike-missing-parent-{Guid.NewGuid():N}";
        var path = $"{missingParent}/notes.txt";

        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("v1"));

        try
        {
            var created = await raw.Drives[driveId].Items["root"].ItemWithPath(path).Content.PutAsync(stream);
            TestContext.Out.WriteLine($"Phase 1 §1 finding: content PUT to a missing parent auto-vivified it (item id {created?.Id}).");
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError error) when (error.ResponseStatusCode == 404)
        {
            TestContext.Out.WriteLine("Phase 1 §1 finding: content PUT to a missing parent 404'd rather than auto-vivifying it.");
        }
        finally
        {
            // Clean up either way: if auto-vivify happened, this removes the whole
            // tree; if it 404'd, this is a no-op against something never created.
            try
            {
                await raw.Drives[driveId].Items["root"].ItemWithPath(missingParent).DeleteAsync();
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError)
            {
                // Nothing to clean up - expected on the 404 branch.
            }
        }
    }
}
