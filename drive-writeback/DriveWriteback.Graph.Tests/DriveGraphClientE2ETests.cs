using Microsoft.Graph;

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
    /// Phase 1 §1 open question: does PUT .../content honor @microsoft.graph.conflictBehavior
    /// as a query parameter the way POST .../children does? Confirmed by reflecting on the
    /// 6.2.0 assembly that ContentRequestBuilder.PutAsync/ToPutRequestInformation take only
    /// Microsoft.Kiota.Abstractions.DefaultQueryParameters - no typed conflictBehavior
    /// parameter exists - so this probes it the only way available: build the request via
    /// the SDK, then append the raw query string to RequestInformation.URI before sending
    /// through RawClient.RequestAdapter directly, bypassing DriveGraphClient's public
    /// surface entirely (this test exists to inform CreateFileAsync's design, not exercise
    /// a method that uses its result).
    ///
    /// CreateFileAsync (§1) does NOT gate on this test's outcome - it pre-checks existence
    /// via TryGetItemByPathAsync and applies conflict_behavior client-side regardless, which
    /// is correct whether or not Graph honors the query parameter. This is a recorded
    /// observation, not a prerequisite.
    ///
    /// RESOLVED against a live tenant: Graph DOES honor the raw query parameter and 409s.
    /// The first live run of this test failed with an uncaught
    /// Microsoft.Kiota.Abstractions.ApiException ("no error factory is registered for this
    /// code: 409") rather than the expected ODataError - this raw, manually-injected-query-
    /// string request has no 409 error factory registered on its RequestInformation, so Kiota
    /// falls back to the bare ApiException base type instead of the richer ODataError
    /// subtype. The 409 itself is exactly the finding this probe exists to observe; only the
    /// catch clause's exception type was wrong. Catching ApiException (ODataError's own base
    /// type) fixes this while still catching ODataError too, if a future SDK version ever
    /// does register a factory here.
    /// </summary>
    [Test]
    public async Task ContentPut_conflictBehavior_query_parameter_is_an_open_question()
    {
        var raw = _client!.RawClient;
        var driveId = (await raw.Me.Drive.GetAsync())!.Id!;
        var path = $"phase0-spike-conflict-probe-{Guid.NewGuid():N}.txt";

        try
        {
            using (var firstStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("v1")))
            {
                await raw.Drives[driveId].Items["root"].ItemWithPath(path).Content.PutAsync(firstStream);
            }

            var contentBuilder = raw.Drives[driveId].Items["root"].ItemWithPath(path).Content;

            using var secondStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("v2-should-conflict-if-honored"));
            var requestInfo = contentBuilder.ToPutRequestInformation(secondStream);
            requestInfo.URI = new Uri($"{requestInfo.URI}?@microsoft.graph.conflictBehavior=fail");

            try
            {
                await raw.RequestAdapter.SendAsync(requestInfo, Microsoft.Graph.Models.DriveItem.CreateFromDiscriminatorValue);
                TestContext.Out.WriteLine(
                    "Phase 1 §1 finding: PUT .../content did NOT honor @microsoft.graph.conflictBehavior=fail - it overwrote instead of 409ing.");
            }
            catch (Microsoft.Kiota.Abstractions.ApiException error) when (error.ResponseStatusCode == 409)
            {
                TestContext.Out.WriteLine(
                    "Phase 1 §1 finding: PUT .../content DOES honor @microsoft.graph.conflictBehavior=fail - it 409'd as expected.");
            }
        }
        finally
        {
            await raw.Drives[driveId].Items["root"].ItemWithPath(path).DeleteAsync();
        }
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

            // Verify by id, not by re-resolving the nested path - CreateFolderPathAsync's doc
            // comment explains why colon-path addressing of a just-created deep tree isn't
            // reliable enough to trust even for this test's own verification step.
            var resolved = await _client.GetItemByIdAsync(secondCreate!.Id!);
            Assert.That(resolved?.Folder, Is.Not.Null, "Expected the deepest segment to be a folder.");
            Assert.That(resolved?.ParentReference?.Path, Does.EndWith("/a/b"), "Expected 'c' to be nested three levels deep under the spike root, not sitting elsewhere.");
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

    /// <summary>
    /// PRD §11 Phase 0: "Confirm site ID and item ID formats so the path-vs-ID
    /// discriminator in §5 is sound." The property that discriminator actually needs is
    /// narrow: a drive-relative path can always contain '/', so an item id containing '/'
    /// would make "is this string a path or already an id" ambiguous. This test reports
    /// the observed id shape for the record and asserts that discriminating property holds.
    /// </summary>
    [Test]
    public async Task GetItemByPathAsync_returns_an_id_that_never_looks_like_a_path()
    {
        var path = $"phase0-spike-id-{Guid.NewGuid():N}.txt";

        try
        {
            await _client!.UploadTextContentAsync(path, "v1");
            var item = await _client.GetItemByPathAsync(path);

            Assert.That(item?.Id, Is.Not.Null.And.Not.Empty);
            TestContext.Out.WriteLine($"PRD §5 finding: OneDrive item id shape = '{item!.Id}' (length {item.Id!.Length}).");
            Assert.That(item.Id, Does.Not.Contain("/"), "An item id containing '/' would collide with the path-vs-ID discriminator §5 needs.");
        }
        finally
        {
            await _client!.DeleteItemAsync(path);
        }
    }

    /// <summary>
    /// Phase 2: RenameItemAsync against a real, freshly-created file. Renames by id, then
    /// verifies the new name by re-fetching by id (not by re-deriving a colon-path from the
    /// new name - same colon-path-on-a-just-touched-item caution as everywhere else in this
    /// codebase). Self-skips like every other test in this fixture; unrun as of this commit.
    /// </summary>
    [Test]
    public async Task RenameItemAsync_renames_a_freshly_created_file()
    {
        var originalPath = $"phase2-spike-rename-{Guid.NewGuid():N}.txt";
        var newName = $"phase2-spike-renamed-{Guid.NewGuid():N}.txt";

        try
        {
            var created = await _client!.UploadTextContentAsync(originalPath, "v1");

            var renamed = await _client.RenameItemAsync(created!.Id!, newName);
            Assert.That(renamed?.Name, Is.EqualTo(newName));

            var resolved = await _client.GetItemByIdAsync(created.Id!);
            Assert.That(resolved?.Name, Is.EqualTo(newName));
        }
        finally
        {
            // Root-level path, so the item's current name (post-rename) is the deletable path.
            await _client!.DeleteItemAsync(newName);
        }
    }

    /// <summary>
    /// Phase 2: MoveItemAsync against real, freshly-created folders. A single spike root
    /// contains both "source" and "destination" subfolders so one recursive delete of the
    /// root cleans up everything created here, regardless of where the moved file ends up.
    /// Self-skips like every other test in this fixture; unrun as of this commit.
    /// </summary>
    [Test]
    public async Task MoveItemAsync_moves_a_freshly_created_file_into_a_new_folder()
    {
        var spikeRoot = $"phase2-spike-move-{Guid.NewGuid():N}";

        try
        {
            var sourceFolder = await _client!.CreateFolderPathAsync($"{spikeRoot}/source");
            var destinationFolder = await _client.CreateFolderPathAsync($"{spikeRoot}/destination");
            var file = await _client.UploadTextContentAsync($"{spikeRoot}/source/notes.txt", "v1");

            var moved = await _client.MoveItemAsync(file!.Id!, destinationFolder!.Id!);
            Assert.That(moved?.ParentReference?.Id, Is.EqualTo(destinationFolder.Id));

            var resolved = await _client.GetItemByIdAsync(file.Id!);
            Assert.That(resolved?.ParentReference?.Id, Is.EqualTo(destinationFolder.Id), "the move must be visible on a fresh re-fetch, not just in the PATCH response.");
        }
        finally
        {
            await _client!.DeleteItemAsync(spikeRoot);
        }
    }

    /// <summary>
    /// Load-bearing for ResolveDriveIdAsync's fail-closed SharePoint-rejection guard: the
    /// allow-list there only permits driveType "business"/"personal" through. This asserts the
    /// live tenant's own OneDrive actually reports one of those two values when its drive id is
    /// passed explicitly (rather than omitted) - if a real M365 OneDrive ever reported
    /// something else, the guard would wrongly reject every explicit-drive-id call against the
    /// signed-in user's own drive, not just SharePoint document libraries. No SharePoint site
    /// is available in this tenant to exercise the rejection side directly (see CLAUDE.md's
    /// Status section) - this is the positive-path half of that guard's live verification.
    /// </summary>
    [Test]
    public async Task GetItemByIdAsync_with_the_signed_in_user_s_own_drive_id_reports_an_allow_listed_driveType()
    {
        var ownDriveId = (await _client!.RawClient.Me.Drive.GetAsync())!.Id!;
        var ownDrive = await _client.RawClient.Drives[ownDriveId].GetAsync();

        TestContext.Out.WriteLine($"SharePoint-rejection guard finding: own OneDrive reports driveType = '{ownDrive?.DriveType}'.");
        Assert.That(ownDrive?.DriveType, Is.EqualTo("business").Or.EqualTo("personal"),
            "ResolveDriveIdAsync's allow-list only permits these two values - if this fails, the guard would reject the user's own OneDrive too.");

        var path = $"phase2-spike-drive-guard-{Guid.NewGuid():N}.txt";

        try
        {
            var created = await _client.UploadTextContentAsync(path, "v1", driveId: ownDriveId);
            Assert.That(created, Is.Not.Null, "an explicit driveId for the user's own OneDrive must still be allowed through the guard.");
        }
        finally
        {
            await _client.DeleteItemAsync(path, driveId: ownDriveId);
        }
    }

    /// <summary>
    /// PRD §7 idempotency: DeleteItemByIdAsync on an item already deleted must swallow the
    /// 404 and return, not throw. Self-skips like every other test in this fixture.
    /// </summary>
    [Test]
    public async Task DeleteItemByIdAsync_is_idempotent_on_an_already_deleted_item()
    {
        var path = $"phase2-spike-delete-idempotent-{Guid.NewGuid():N}.txt";
        var created = await _client!.UploadTextContentAsync(path, "v1");

        await _client.DeleteItemByIdAsync(created!.Id!);

        Assert.That(
            async () => await _client.DeleteItemByIdAsync(created.Id!),
            Throws.Nothing,
            "deleting an already-deleted item must succeed, not throw.");
    }
}
