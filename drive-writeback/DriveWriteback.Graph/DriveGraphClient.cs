using System.Collections.Concurrent;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item;
using Microsoft.Graph.Drives.Item.Items.Item.Children;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using DriveWriteback.Graph.Auth;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("DriveWriteback.Graph.Tests")]

namespace DriveWriteback.Graph;

/// <summary>
/// Thin wrapper over the Graph SDK for the full drive-writeback tool surface: delegated
/// auth, folder creation, content upload/replacement, rename/move/delete, and the
/// SharePoint-drive rejection guard, against a self-registered Entra app. Started as the
/// Phase 0 spike client and grew into this as Phase 1/2 landed - see
/// docs/archive/PRD-drive-write.md §11.
///
/// Every method takes an optional driveId, mirroring the PRD §5 addressing model: omitted,
/// it resolves the signed-in user's own OneDrive (GET /me/drive); supplied, it targets
/// that drive directly (e.g. a SharePoint document library's drive id).
///
/// The Graph SDK's generated RootRequestBuilder (under Me.Drive.Root) only exposes
/// GetAsync/Content - no Children, no ItemWithPath - so folder/file operations route
/// through Drives[driveId].Items["root"] instead, using "root" as Graph's documented
/// alias item id. That builder type does support ItemWithPath (a generated extension
/// method) and Children, which is what mkdir -p and path-based addressing need.
/// </summary>
public sealed class DriveGraphClient(GraphServiceClient client)
{
    /// <summary>
    /// Test-only escape hatch for probing Graph SDK behavior this class doesn't wrap yet -
    /// see the conflictBehavior-on-content-PUT probe in DriveGraphClientE2ETests, which
    /// needs to build a request this class has no public method for. Internal, not public:
    /// nothing outside the InternalsVisibleTo'd test project should reach through this.
    /// </summary>
    internal GraphServiceClient RawClient => client;

    /// <summary>
    /// Drive ids already confirmed (via ResolveDriveIdAsync) to be a OneDrive, not a SharePoint
    /// document library - avoids a repeat GET /drives/{id} on every call in the same chain (e.g.
    /// CreateFileAsync's parent-check/target-check/rename-collision-loop all resolve the same
    /// driveId). This class is registered as a DI singleton (Program.cs), so the set is
    /// concurrent and lives for the process lifetime - a drive's type never changes, so caching
    /// it forever is safe.
    /// </summary>
    private readonly ConcurrentDictionary<string, byte> validatedOneDriveIds = new();

    /// <summary>
    /// Requires the Entra app to be registered as a public client with
    /// "http://localhost" listed under Mobile and desktop redirect URIs - a
    /// confidential/web registration will fail the redirect-URI check at sign-in.
    /// </summary>
    public static DriveGraphClient CreateWithInteractiveBrowserAuth(
        string tenantId,
        string clientId,
        IEnumerable<string> scopes)
    {
        var credential = new InteractiveBrowserCredential(
            new InteractiveBrowserCredentialOptions
            {
                TenantId = tenantId,
                ClientId = clientId,
                RedirectUri = new Uri("http://localhost"),
            });

        var client = new GraphServiceClient(credential, scopes);

        return new DriveGraphClient(client);
    }

    /// <summary>
    /// Non-interactive counterpart to CreateWithInteractiveBrowserAuth, for a deployed
    /// service that can't pop a browser - silently redeems a refresh token cached in
    /// refreshTokenStore instead. The refresh token itself must be seeded once via a
    /// separate interactive bootstrap step (see DriveWriteback.Bootstrap).
    /// </summary>
    public static DriveGraphClient CreateWithSilentRefreshAuth(
        string tenantId,
        string clientId,
        IEnumerable<string> scopes,
        IRefreshTokenStore refreshTokenStore)
    {
        var scopeArray = scopes as string[] ?? [.. scopes];
        var tokenEndpointClient = new GraphTokenEndpointClient(tenantId, clientId);
        var credential = new SilentGraphCredential(tokenEndpointClient, refreshTokenStore, scopeArray);
        var client = new GraphServiceClient(credential, scopeArray);

        return new DriveGraphClient(client);
    }

    /// <summary>
    /// Creates a single child folder under the given drive-relative parent path.
    /// Conflict behavior "fail" - callers implementing mkdir -p treat the resulting
    /// 409 (surfaced as a DriveItemAlreadyExistsException) as "already exists, continue."
    /// Pass null or "" for parentPath to create directly under the drive root.
    ///
    /// Safe when parentPath refers to a stable, already-existing folder. NOT safe to chain
    /// against a parent this same call chain just created a moment ago - see
    /// CreateFolderPathAsync's doc comment for why, and use that method instead for
    /// mkdir-p-style nested creation.
    /// </summary>
    public async Task<DriveItem?> CreateFolderSegmentAsync(
        string? parentPath,
        string folderName,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);
        var parent = ResolveRootItem(resolvedDriveId, parentPath);
        var newFolder = new DriveItem
        {
            Name = folderName,
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
            {
                ["@microsoft.graph.conflictBehavior"] = "fail",
            },
        };

        try
        {
            return await parent.Children.PostAsync(newFolder, cancellationToken: cancellationToken);
        }
        catch (ODataError error) when (error.ResponseStatusCode == 409)
        {
            throw new DriveItemAlreadyExistsException(folderName, error);
        }
    }

    /// <summary>
    /// mkdir -p: walks path segments from the drive root, creating each missing one.
    /// Idempotent - an existing full path returns success rather than an error, per
    /// PRD §4.4's baseline approach (iterative per-segment create, treating 409 as
    /// "already exists, continue"). This is the Phase 0 spike artifact for that section.
    ///
    /// Chains by item id, not by re-deriving a colon-path string from segment names.
    /// Live E2E testing found that colon-path addressing of a
    /// folder immediately after creating it is not reliable: Graph's path-resolution index
    /// can lag behind the item actually existing, and the failure mode observed wasn't a
    /// clean 404 - a /children POST against an unresolved colon-path silently landed at the
    /// drive root instead of erroring, producing folders that looked created but weren't
    /// nested where intended (confirmed both via a direct OneDrive-web check and by logging
    /// each created item's own ParentReference.Path). Addressing by the id Graph just
    /// returned needs no path resolution at all and sidesteps this entirely.
    /// See docs/archive/PRD-drive-write.md §11/§12 Q3.
    /// </summary>
    public async Task<DriveItem?> CreateFolderPathAsync(
        string fullPath,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var segments = fullPath.Split('/', StringSplitOptions.RemoveEmptyEntries);

        if (segments.Length == 0)
            throw new ArgumentException("fullPath must contain at least one segment.", nameof(fullPath));

        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);
        string? parentId = null;
        DriveItem? current = null;

        foreach (var segment in segments)
        {
            try
            {
                current = await CreateChildByParentIdAsync(resolvedDriveId, parentId, segment, cancellationToken);
            }
            catch (DriveItemAlreadyExistsException)
            {
                current = await ResolveExistingChildByParentIdAsync(resolvedDriveId, parentId, segment, cancellationToken);
            }

            parentId = current?.Id
                ?? throw new InvalidOperationException($"Could not resolve an id for segment '{segment}' after create or already-exists lookup.");
        }

        return current;
    }

    /// <summary>
    /// Creates a child folder directly under a known parent item id (or the drive root when
    /// parentId is null). Id-based, not path-based - see CreateFolderPathAsync's doc comment.
    /// </summary>
    private async Task<DriveItem?> CreateChildByParentIdAsync(
        string driveId,
        string? parentId,
        string folderName,
        CancellationToken cancellationToken)
    {
        var newFolder = new DriveItem
        {
            Name = folderName,
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
            {
                ["@microsoft.graph.conflictBehavior"] = "fail",
            },
        };
        var children = ResolveChildrenByParentId(driveId, parentId);

        try
        {
            return await children.PostAsync(newFolder, cancellationToken: cancellationToken);
        }
        catch (ODataError error) when (error.ResponseStatusCode == 409)
        {
            throw new DriveItemAlreadyExistsException(folderName, error);
        }
    }

    /// <summary>
    /// Resolves an already-existing child by listing its parent's children (by id) and
    /// filtering by name - id-based, same reasoning as CreateChildByParentIdAsync.
    /// </summary>
    private async Task<DriveItem?> ResolveExistingChildByParentIdAsync(
        string driveId,
        string? parentId,
        string childName,
        CancellationToken cancellationToken)
    {
        var children = ResolveChildrenByParentId(driveId, parentId);
        var escapedName = childName.Replace("'", "''");

        var result = await children.GetAsync(
            requestConfiguration => requestConfiguration.QueryParameters.Filter = $"name eq '{escapedName}'",
            cancellationToken: cancellationToken);

        return result?.Value?.FirstOrDefault();
    }

    private ChildrenRequestBuilder ResolveChildrenByParentId(string driveId, string? parentId) =>
        string.IsNullOrEmpty(parentId)
            ? client.Drives[driveId].Items["root"].Children
            : client.Drives[driveId].Items[parentId].Children;

    /// <summary>
    /// Fetches an item by its stable id rather than a path - use this instead of
    /// GetItemByPathAsync to verify an item that was just created moments ago in the same
    /// call chain (see CreateFolderPathAsync's doc comment for why path addressing of a
    /// freshly-created item isn't reliable).
    /// </summary>
    public async Task<DriveItem?> GetItemByIdAsync(
        string itemId,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);

        return await client.Drives[resolvedDriveId].Items[itemId].GetAsync(cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Wraps GetItemByIdAsync, returning null on a 404 instead of throwing - the same
    /// existence-check shape TryGetItemByPathAsync provides for paths. delete_item's confirm
    /// step needs this to treat an already-deleted item as success (PRD §7 idempotency)
    /// rather than letting the 404 propagate as an error.
    /// </summary>
    public async Task<DriveItem?> TryGetItemByIdAsync(
        string itemId,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            return await GetItemByIdAsync(itemId, driveId, cancellationToken);
        }
        catch (ODataError error) when (error.ResponseStatusCode == 404)
        {
            return null;
        }
    }

    /// <summary>
    /// The path-vs-ID discriminator wrapper get_item needs (PRD §5): decides which of
    /// GetItemByIdAsync/GetItemByPathAsync to call using DrivePath.LooksLikeItemId, and
    /// reports which interpretation it used so a misclassification is visible to the
    /// caller rather than silent (see DrivePath.LooksLikeItemId's doc comment).
    /// </summary>
    public async Task<ItemResolution> GetItemAsync(
        string pathOrId,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        if (DrivePath.LooksLikeItemId(pathOrId))
            return new ItemResolution(await GetItemByIdAsync(pathOrId, driveId, cancellationToken), ResolvedAsId: true);

        return new ItemResolution(await GetItemByPathAsync(pathOrId, driveId, cancellationToken), ResolvedAsId: false);
    }

    /// <summary>
    /// Uploads small text content (Graph's simple-upload path, at or below 4 MB) to the
    /// given drive-relative path. Fails outright if something already exists there -
    /// this spike doesn't exercise conflictBehavior on upload, only on folder creation.
    /// </summary>
    public async Task<DriveItem?> UploadTextContentAsync(
        string path,
        string content,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(content));

        return await ResolvePathItem(resolvedDriveId, path)
            .Content
            .PutAsync(stream, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Fetches an item's metadata, including both eTag and cTag - the Phase 0 spike
    /// this exists for is determining which one Graph actually honors for If-Match
    /// on a subsequent content PUT (PRD §12 Q5).
    /// </summary>
    public async Task<DriveItem?> GetItemByPathAsync(
        string path,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);

        return await ResolvePathItem(resolvedDriveId, path).GetAsync(cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Wraps GetItemByPathAsync, returning null on a 404 instead of throwing - the
    /// existence check create_file's conflict_behavior handling and ResolveFolderPathAsync
    /// both need.
    /// </summary>
    public async Task<DriveItem?> TryGetItemByPathAsync(
        string path,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            return await GetItemByPathAsync(path, driveId, cancellationToken);
        }
        catch (ODataError error) when (error.ResponseStatusCode == 404)
        {
            return null;
        }
    }

    /// <summary>
    /// Replaces an existing file's content, requiring the caller to supply an If-Match
    /// tag value (eTag or cTag - the spike test tries both). Surfaces a 412 distinctly
    /// so the test can assert on which tag Graph actually enforced against.
    /// </summary>
    public async Task<DriveItem?> ReplaceTextContentAsync(
        string path,
        string content,
        string ifMatchTag,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(content));

        try
        {
            return await ResolvePathItem(resolvedDriveId, path)
                .Content
                .PutAsync(
                    stream,
                    requestConfiguration => requestConfiguration.Headers.TryAdd("If-Match", ifMatchTag),
                    cancellationToken);
        }
        catch (ODataError error) when (error.ResponseStatusCode == 412)
        {
            throw new DriveItemConcurrencyException(path, ifMatchTag, error);
        }
    }

    public async Task DeleteItemAsync(
        string path,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);

        await ResolvePathItem(resolvedDriveId, path).DeleteAsync(cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Read-only counterpart to CreateFolderPathAsync's mkdir-p walk: resolves as far as
    /// existing folders go and reports which trailing segments are missing, without ever
    /// POSTing. This is what dry-run mode calls instead of actually creating anything.
    /// Reuses the same id-chained ResolveExistingChildByParentIdAsync CreateFolderPathAsync
    /// uses, for the same colon-path-is-unreliable reason documented there.
    /// </summary>
    public async Task<FolderPathResolution> ResolveFolderPathAsync(
        string fullPath,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var normalized = DrivePath.Normalize(fullPath);

        if (normalized.Length == 0)
            return new FolderPathResolution(null, []);

        var segments = normalized.Split('/');
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);
        string? parentId = null;
        DriveItem? deepestExisting = null;

        for (var i = 0; i < segments.Length; i++)
        {
            var found = await ResolveExistingChildByParentIdAsync(resolvedDriveId, parentId, segments[i], cancellationToken);

            if (found is null)
                return new FolderPathResolution(deepestExisting, segments[i..]);

            deepestExisting = found;
            parentId = found.Id;
        }

        return new FolderPathResolution(deepestExisting, []);
    }

    /// <summary>
    /// Creates a file with inline text content. Parents are NOT auto-created (PRD §4.4) -
    /// callers wanting mkdir-p call CreateFolderPathAsync first; this throws
    /// DriveParentNotFoundException if the parent path doesn't resolve.
    ///
    /// conflict_behavior is enforced client-side - pre-check via TryGetItemByPathAsync, then
    /// "fail" throws, "rename" appends a numeric suffix, "replace" PUTs unconditionally -
    /// rather than via Graph's @microsoft.graph.conflictBehavior query parameter on the
    /// simple-upload PUT endpoint. See
    /// DriveGraphClientE2ETests.ContentPut_conflictBehavior_query_parameter_is_an_open_question:
    /// the SDK exposes no typed support for that parameter, and whether Graph honors it as a
    /// raw query string is an unrun probe, not a confirmed fact - this implementation doesn't
    /// depend on the answer either way. The client-side "fail"/"rename" pre-check is
    /// TOCTOU-racy (a concurrent writer between the check and the PUT could still be
    /// silently overwritten); tighten to a server-enforced conflictBehavior if that probe
    /// later confirms Graph honors it.
    /// </summary>
    public async Task<DriveItem?> CreateFileAsync(
        string path,
        string content,
        string conflictBehavior = "fail",
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        if (conflictBehavior is not ("fail" or "rename" or "replace"))
            throw new ArgumentException($"conflictBehavior must be 'fail', 'rename', or 'replace' - got '{conflictBehavior}'.", nameof(conflictBehavior));

        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);
        var (parentPath, name) = DrivePath.SplitParent(path);

        if (parentPath.Length > 0 && await TryGetItemByPathAsync(parentPath, resolvedDriveId, cancellationToken) is null)
            throw new DriveParentNotFoundException(parentPath);

        var targetPath = path;

        if (conflictBehavior is "fail" or "rename")
        {
            var existing = await TryGetItemByPathAsync(path, resolvedDriveId, cancellationToken);

            if (existing is not null)
            {
                if (conflictBehavior == "fail")
                    throw new DriveItemAlreadyExistsException(name);

                targetPath = await ResolveNonCollidingPathAsync(parentPath, name, resolvedDriveId, cancellationToken);
            }
        }

        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(content));

        return await ResolvePathItem(resolvedDriveId, targetPath)
            .Content
            .PutAsync(stream, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Appends a numeric " (n)" suffix (before the extension) until a path that doesn't
    /// already exist is found - the client-side implementation of conflict_behavior:
    /// "rename". Capped rather than unbounded: a runaway loop here would mean something is
    /// wrong with the existence check, not that 1000 same-named files genuinely exist.
    /// </summary>
    private async Task<string> ResolveNonCollidingPathAsync(
        string parentPath,
        string name,
        string driveId,
        CancellationToken cancellationToken)
    {
        var dotIndex = name.LastIndexOf('.');
        var stem = dotIndex < 0 ? name : name[..dotIndex];
        var extension = dotIndex < 0 ? "" : name[dotIndex..];

        for (var suffix = 1; suffix <= 1000; suffix++)
        {
            var candidateName = $"{stem} ({suffix}){extension}";
            var candidatePath = parentPath.Length == 0 ? candidateName : $"{parentPath}/{candidateName}";

            if (await TryGetItemByPathAsync(candidatePath, driveId, cancellationToken) is null)
                return candidatePath;
        }

        throw new InvalidOperationException($"Could not find a non-colliding name for '{parentPath}/{name}' after 1000 attempts.");
    }

    /// <summary>
    /// Replaces an existing file's content by item id, requiring an If-Match tag - the
    /// path_or_id-aware, id-preferring promotion of ReplaceTextContentAsync that
    /// update_file_content actually calls. The service layer resolves path_or_id to an id
    /// via GetItemAsync first; this method only ever acts by id, for the same
    /// colon-path-is-unreliable-on-a-just-touched-item reason documented on
    /// CreateFolderPathAsync.
    /// </summary>
    public async Task<DriveItem?> ReplaceContentByIdAsync(
        string itemId,
        string content,
        string ifMatchTag,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(content));

        try
        {
            return await client.Drives[resolvedDriveId].Items[itemId]
                .Content
                .PutAsync(
                    stream,
                    requestConfiguration => requestConfiguration.Headers.TryAdd("If-Match", ifMatchTag),
                    cancellationToken);
        }
        catch (ODataError error) when (error.ResponseStatusCode == 412)
        {
            throw new DriveItemConcurrencyException(itemId, ifMatchTag, error);
        }
    }

    /// <summary>
    /// Shared PATCH primitive backing RenameItemAsync and MoveItemAsync (PRD §4.4: "Both are
    /// PATCH /drives/{drive-id}/items/{id} - rename sets name, move sets parentReference.id").
    /// Id-based only, like every other mutating method here.
    /// </summary>
    private async Task<DriveItem?> PatchItemAsync(
        string driveId,
        string itemId,
        DriveItem patch,
        CancellationToken cancellationToken)
    {
        try
        {
            return await client.Drives[driveId].Items[itemId].PatchAsync(patch, cancellationToken: cancellationToken);
        }
        catch (ODataError error) when (error.ResponseStatusCode == 409)
        {
            throw new DriveItemAlreadyExistsException(patch.Name ?? itemId, error);
        }
    }

    /// <summary>
    /// Renames an item in place (PATCH { name }). Id-based - the service layer resolves
    /// path_or_id to an id via GetItemAsync first.
    /// </summary>
    public async Task<DriveItem?> RenameItemAsync(
        string itemId,
        string newName,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);

        return await PatchItemAsync(resolvedDriveId, itemId, new DriveItem { Name = newName }, cancellationToken);
    }

    /// <summary>
    /// Moves an item to a new parent within the same drive (PATCH { parentReference.id }).
    /// PRD §4.4: move_item is same-drive-only - cross-drive moves require copy+delete, out of
    /// scope (PRD §3, §11 Phase 4). Graph's own 404 on a foreign parent id is the primary
    /// enforcement of that; the ParentReference.DriveId check below is defense-in-depth for the
    /// narrower case where a resolved destination item reports a different owning drive than
    /// the one this call was made under (e.g. a remote/shared item alias).
    /// </summary>
    public async Task<DriveItem?> MoveItemAsync(
        string itemId,
        string newParentId,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);
        var destinationParent = await client.Drives[resolvedDriveId].Items[newParentId]
            .GetAsync(cancellationToken: cancellationToken)
            ?? throw new DriveParentNotFoundException(newParentId);

        if (destinationParent.ParentReference?.DriveId is { } actualDriveId && actualDriveId != resolvedDriveId)
            throw new DriveCrossDriveMoveException(resolvedDriveId, actualDriveId);

        return await PatchItemAsync(
            resolvedDriveId,
            itemId,
            new DriveItem { ParentReference = new ItemReference { Id = newParentId } },
            cancellationToken);
    }

    /// <summary>
    /// Deletes an item by id, moving it to the recycle bin (PRD §7/§9: never permanent delete).
    /// A 404 is swallowed rather than thrown - PRD §7 idempotency requires delete_item on an
    /// already-deleted item to return success, not error. The recursive/non-empty-folder guard
    /// (PRD §4.4) is enforced one layer up in DriveItemDeletionService against the already-
    /// resolved item's Folder.ChildCount, not here - this method stays a dumb Graph wrapper.
    /// </summary>
    public async Task DeleteItemByIdAsync(
        string itemId,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var resolvedDriveId = await ResolveDriveIdAsync(driveId, cancellationToken);

        try
        {
            await client.Drives[resolvedDriveId].Items[itemId].DeleteAsync(cancellationToken: cancellationToken);
        }
        catch (ODataError error) when (error.ResponseStatusCode == 404)
        {
        }
    }

    private async Task<string> ResolveOwnDriveIdAsync(CancellationToken cancellationToken)
    {
        var drive = await client.Me.Drive.GetAsync(cancellationToken: cancellationToken);

        return drive?.Id ?? throw new InvalidOperationException("Could not resolve the signed-in user's OneDrive id.");
    }

    /// <summary>
    /// Resolves the effective drive id every method above acts against, and is this server's
    /// one enforcement point for staying OneDrive-only: omitted, the signed-in user's own
    /// OneDrive (GET /me/drive can never resolve to a SharePoint site drive, so no further
    /// check is needed); supplied, the given id is validated by fetching its Drive resource and
    /// checking driveType. SharePoint document libraries report "documentLibrary" - this fails
    /// closed on that and on anything unrecognized (a missing/null driveType, or a value this
    /// tenant hasn't been observed to return), allowing only "business"/"personal" through, so
    /// an unanticipated Graph response can't silently bypass the guard. This is a deliberate
    /// product decision, not a Phase-3-readiness gap: SharePoint hardening (checkout detection,
    /// required-column draft-state, version reporting - PRD §11 Phase 3) was never built, so
    /// this deployment rejects SharePoint drives outright rather than operating against them
    /// unsafely. "Shared with me" items in another user's OneDrive still work - those report
    /// driveType "business" too, same as the signed-in user's own drive.
    /// Validated ids are cached in validatedOneDriveIds so a single call chain that touches the
    /// same driveId repeatedly (e.g. CreateFileAsync's parent/target checks and its
    /// rename-collision loop) only pays for one GET /drives/{id}.
    /// </summary>
    private async Task<string> ResolveDriveIdAsync(string? driveId, CancellationToken cancellationToken)
    {
        if (driveId is null)
            return await ResolveOwnDriveIdAsync(cancellationToken);

        if (validatedOneDriveIds.ContainsKey(driveId))
            return driveId;

        var drive = await client.Drives[driveId].GetAsync(cancellationToken: cancellationToken);

        if (drive?.DriveType is not ("business" or "personal"))
            throw new SharePointDriveNotSupportedException(driveId, drive?.DriveType);

        validatedOneDriveIds[driveId] = 0;

        return driveId;
    }

    private DriveItemItemRequestBuilder ResolveRootItem(string driveId, string? path) =>
        string.IsNullOrEmpty(path)
            ? client.Drives[driveId].Items["root"]
            : client.Drives[driveId].Items["root"].ItemWithPath(path);

    private CustomDriveItemItemRequestBuilder ResolvePathItem(string driveId, string path) =>
        client.Drives[driveId].Items["root"].ItemWithPath(path);
}

/// <summary>
/// Surfaced when a folder-segment create hits Graph's 409 for an existing name (inner set),
/// or when create_file's client-side conflict_behavior: "fail" pre-check finds an existing
/// target before ever calling Graph (inner null - there's no ODataError to attach, since
/// no request was made).
/// </summary>
public sealed class DriveItemAlreadyExistsException(string itemName, ODataError? inner = null)
    : Exception($"An item named '{itemName}' already exists at this location.", inner)
{
    public string ItemName { get; } = itemName;
}

/// <summary>
/// Result of a read-only mkdir-p walk (see DriveGraphClient.ResolveFolderPathAsync):
/// DeepestExisting is the last segment found to already exist (null if even the first
/// segment is missing); MissingSegments is every segment from the first missing one
/// onward, in order.
/// </summary>
public sealed record FolderPathResolution(DriveItem? DeepestExisting, IReadOnlyList<string> MissingSegments);

/// <summary>
/// Result of DriveGraphClient.GetItemAsync's path-vs-ID discriminator: Item is the
/// resolved item (null if not found), ResolvedAsId records which interpretation of the
/// input string was used.
/// </summary>
public sealed record ItemResolution(DriveItem? Item, bool ResolvedAsId);

/// <summary>
/// Surfaced when a content replacement's If-Match tag doesn't match Graph's current
/// value - the distinct, actionable 412 the PRD (§4.4) requires callers to see.
/// </summary>
public sealed class DriveItemConcurrencyException(string path, string attemptedTag, ODataError inner)
    : Exception($"If-Match '{attemptedTag}' did not match the current state of '{path}'.", inner)
{
    public string Path { get; } = path;
    public string AttemptedTag { get; } = attemptedTag;
}

/// <summary>
/// Surfaced by create_file/create_folder when the target's parent path does not exist -
/// per PRD §4.4, parents are not auto-created, so a typo'd path fails loudly rather than
/// silently materializing a folder tree.
/// </summary>
public sealed class DriveParentNotFoundException(string parentPath)
    : Exception($"Parent path '{parentPath}' does not exist. Parents are not auto-created - call create_folder first.")
{
    public string ParentPath { get; } = parentPath;
}

/// <summary>
/// Surfaced when create_file's content exceeds the configured MaxContentBytes cap (PRD §7,
/// §9 item 5) - enforced before the Graph call, not left to Graph's own 4 MB simple-upload
/// ceiling.
/// </summary>
public sealed class DriveContentTooLargeException(long actualBytes, long maxBytes)
    : Exception($"Content is {actualBytes} bytes, exceeding the {maxBytes}-byte limit.")
{
    public long ActualBytes { get; } = actualBytes;
    public long MaxBytes { get; } = maxBytes;
}

/// <summary>
/// Surfaced by move_item's defense-in-depth check when a resolved destination parent reports
/// a ParentReference.DriveId different from the drive context the call was made under - Graph's
/// own PATCH would 404 in the ordinary case; this catches the narrower remote/shared-item
/// aliasing case before ever calling Graph. Cross-drive move itself is out of scope (PRD §3,
/// §11 Phase 4) - this is a clear error, not an attempted copy+delete fallback.
/// </summary>
public sealed class DriveCrossDriveMoveException(string expectedDriveId, string actualDriveId)
    : Exception($"Destination resolves to drive '{actualDriveId}', not the expected '{expectedDriveId}'. Cross-drive moves are not supported.")
{
    public string ExpectedDriveId { get; } = expectedDriveId;
    public string ActualDriveId { get; } = actualDriveId;
}

/// <summary>
/// Surfaced by ResolveDriveIdAsync when an explicitly-supplied drive_id resolves to a
/// SharePoint document library (driveType "documentLibrary") or to anything else that isn't
/// a recognized OneDrive ("business"/"personal"). This deployment deliberately does not
/// support SharePoint document libraries - Phase 3's hardening (checkout detection,
/// required-column draft-state, version reporting; PRD §11) was never built, so writing to a
/// SharePoint library here would hit failure modes this server can't yet detect or explain.
/// </summary>
public sealed class SharePointDriveNotSupportedException(string driveId, string? actualDriveType)
    : Exception($"Drive '{driveId}' is not a supported OneDrive (reported driveType: '{actualDriveType ?? "unknown"}'). " +
        "This server only supports OneDrive; SharePoint document libraries are not supported.")
{
    public string DriveId { get; } = driveId;
    public string? ActualDriveType { get; } = actualDriveType;
}

/// <summary>
/// Surfaced when delete_item targets a non-empty folder without recursive=true (PRD §4.4).
/// </summary>
public sealed class DriveFolderNotEmptyException(string itemName, int childCount)
    : Exception($"'{itemName}' contains {childCount} item(s). Pass recursive=true to delete a non-empty folder.")
{
    public string ItemName { get; } = itemName;
    public int ChildCount { get; } = childCount;
}

/// <summary>
/// Surfaced when update_file_content/rename_item/move_item's path_or_id resolves to nothing -
/// Phase 1's tools never needed this distinct case (create tools check existence pre-write or
/// rely on Graph's own 404), but the mutation tools all resolve their target via GetItemAsync
/// first and need a clear signal when that resolution comes back empty.
/// </summary>
public sealed class ItemNotFoundException(string pathOrId)
    : Exception($"No item found at '{pathOrId}'.")
{
    public string PathOrId { get; } = pathOrId;
}

/// <summary>
/// Surfaced by GetItemByIdAsync/GetItemByPathAsync when a caller-supplied not_modified_since
/// value predates the resolved item's Graph LastModifiedDateTime (PRD-etag-enforcement.md §4.1)
/// - the item changed after the caller's own read, so handing back a fresh if_match here would
/// let a stale edit silently overwrite that change. Client-side only; Graph has no server-side
/// If-Unmodified-Since equivalent for driveItem.
/// </summary>
public sealed class ItemModifiedSinceReadException(string pathOrId, DateTimeOffset notModifiedSince, DateTimeOffset lastModifiedDateTime)
    : Exception($"'{pathOrId}' was modified at {lastModifiedDateTime:O}, after the supplied not_modified_since of {notModifiedSince:O}. " +
        "Re-read the file's content and resolve any conflict before retrying.")
{
    public string PathOrId { get; } = pathOrId;
    public DateTimeOffset NotModifiedSince { get; } = notModifiedSince;
    public DateTimeOffset LastModifiedDateTime { get; } = lastModifiedDateTime;
}
