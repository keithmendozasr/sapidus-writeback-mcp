using System.Text;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Graph;
using DriveWriteback.Graph.Confirmation;
using DriveWriteback.Graph.Writes;
using Sapidus.Writeback.Shared.Confirmation;

namespace DriveWriteback.Graph.Tests.Writes;

/// <summary>
/// True end-to-end tier for the Phase 1 write surface: real Graph API, real OneDrive.
/// Requires the "Drive Writeback MCP" Entra app (see ../DEPLOYMENT.md) and self-skips via
/// Assert.Ignore when the environment variables below aren't set. Complements
/// DriveGraphClientE2ETests, which exercises DriveGraphClient directly - this fixture goes
/// through DriveWriteService, the layer create_folder/create_file's tool classes actually
/// call.
///
/// All tests in this fixture have now been run against a live tenant and pass - see
/// docs/archive/PRD-drive-write.md §11 and ../CLAUDE.md's Status section for the findings
/// these runs resolved.
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

    /// <summary>
    /// PRD §9 item 8: dry-run mode must validate and resolve but never mutate. Runs both
    /// create_folder and create_file through a dry-run DriveWriteService against the live
    /// tenant, then confirms via TryGetItemByPathAsync (a plain read, bypassing the
    /// path-vs-id discriminator since these are known paths, not ids) that neither actually
    /// got created.
    /// </summary>
    [Test]
    public async Task DryRun_makes_zero_mutations_against_the_live_tenant()
    {
        var dryRunService = CreateService(dryRun: true);
        var folderPath = $"phase1-spike-dryrun-folder-{Guid.NewGuid():N}";
        var filePath = $"phase1-spike-dryrun-file-{Guid.NewGuid():N}.txt";

        var folderResult = await dryRunService.CreateFolderAsync(folderPath);
        var fileResult = await dryRunService.CreateFileAsync(filePath, "content");

        Assert.Multiple(() =>
        {
            Assert.That(folderResult.DryRun, Is.True);
            Assert.That(fileResult.DryRun, Is.True);
        });

        var folderCheck = await _client!.TryGetItemByPathAsync(folderPath);
        var fileCheck = await _client.TryGetItemByPathAsync(filePath);

        Assert.Multiple(() =>
        {
            Assert.That(folderCheck, Is.Null, "dry-run create_folder must not actually create anything.");
            Assert.That(fileCheck, Is.Null, "dry-run create_file must not actually create anything.");
        });
    }

    /// <summary>
    /// Phase 2: update_file_content's mandatory if_match against a real file. Dry-run must
    /// resolve the target and report the pre-update eTag without writing; real mode must
    /// then replace the content and return a different post-update eTag. Self-skips like
    /// every other test in this fixture; unrun as of this commit.
    /// </summary>
    [Test]
    public async Task UpdateFileContentAsync_dry_run_then_real_mode_round_trip()
    {
        var path = $"phase2-spike-update-content-{Guid.NewGuid():N}.txt";

        try
        {
            var created = await _client!.UploadTextContentAsync(path, "v1");
            var originalETag = created!.ETag!;

            var dryRunService = CreateService(dryRun: true);
            var dryRunResult = await dryRunService.UpdateFileContentAsync(path, "v2-should-not-write", originalETag);

            Assert.Multiple(() =>
            {
                Assert.That(dryRunResult.DryRun, Is.True);
                Assert.That(dryRunResult.Item?.ETag, Is.EqualTo(originalETag), "dry-run must report the pre-update item, not mutate it.");
            });

            var unchanged = await _client.GetItemByPathAsync(path);
            Assert.That(unchanged?.ETag, Is.EqualTo(originalETag), "dry-run update_file_content must not actually write anything.");

            var realService = CreateService(dryRun: false);
            var realResult = await realService.UpdateFileContentAsync(path, "v2", originalETag);

            Assert.Multiple(() =>
            {
                Assert.That(realResult.DryRun, Is.False);
                Assert.That(realResult.Item?.ETag, Is.Not.EqualTo(originalETag), "a real content replacement must produce a new eTag.");
            });
        }
        finally
        {
            await _client!.DeleteItemAsync(path);
        }
    }

    /// <summary>
    /// Phase 2: delete_item's two-call confirmation flow against a real file, end to end
    /// through DriveItemDeletionService (not just DriveGraphClient directly) - preview issues
    /// a token without deleting, confirm then actually deletes. The signing key here is
    /// purely local to this test process, unrelated to the deployed server's Key Vault
    /// secret. Self-skips like every other test in this fixture; unrun as of this commit.
    /// </summary>
    [Test]
    public async Task DriveItemDeletionService_preview_then_confirm_round_trip_deletes_a_real_file()
    {
        var path = $"phase2-spike-delete-confirm-{Guid.NewGuid():N}.txt";
        var created = await _client!.UploadTextContentAsync(path, "v1");

        var tokenService = new ConfirmationTokenService(Encoding.UTF8.GetBytes("e2e-local-signing-key"));
        var deletionService = new DriveItemDeletionService(_client!, CreateService(dryRun: false), tokenService);

        var pending = await deletionService.RequestDeletionAsync(created!.Id!, created.Name!, recursive: false);
        Assert.That(pending.ItemId, Is.EqualTo(created.Id));

        var stillThere = await _client.TryGetItemByPathAsync(path);
        Assert.That(stillThere, Is.Not.Null, "the preview call must not delete anything.");

        var confirmed = await deletionService.ConfirmDeletionAsync(
            pending.ItemId, pending.ConfirmationToken, created.Name!, recursive: false);
        Assert.That(confirmed, Is.EqualTo(ConfirmedItemDeletion.Deleted));

        var afterDelete = await _client.TryGetItemByPathAsync(path);
        Assert.That(afterDelete, Is.Null, "the confirm call must actually delete the item.");
    }
}
