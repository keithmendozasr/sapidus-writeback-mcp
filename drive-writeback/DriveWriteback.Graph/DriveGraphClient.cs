using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item;
using Microsoft.Graph.Drives.Item.Items.Item.Children;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("DriveWriteback.Graph.Tests")]

namespace DriveWriteback.Graph;

/// <summary>
/// Thin wrapper over the Graph SDK for the Phase 0 spike: proves delegated auth plus
/// folder creation, content upload, and If-Match content replacement work against a
/// self-registered Entra app. Not the full create_file/create_folder/update_file_content
/// tool surface - that lands in Phase 1-2 once these assumptions are verified.
/// See docs/active/PRD-drive-write.md §11.
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
        var resolvedDriveId = driveId ?? await ResolveOwnDriveIdAsync(cancellationToken);
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
    /// See docs/active/PRD-drive-write.md §11/§12 Q3.
    /// </summary>
    public async Task<DriveItem?> CreateFolderPathAsync(
        string fullPath,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var segments = fullPath.Split('/', StringSplitOptions.RemoveEmptyEntries);

        if (segments.Length == 0)
            throw new ArgumentException("fullPath must contain at least one segment.", nameof(fullPath));

        var resolvedDriveId = driveId ?? await ResolveOwnDriveIdAsync(cancellationToken);
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
        var resolvedDriveId = driveId ?? await ResolveOwnDriveIdAsync(cancellationToken);

        return await client.Drives[resolvedDriveId].Items[itemId].GetAsync(cancellationToken: cancellationToken);
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
        var resolvedDriveId = driveId ?? await ResolveOwnDriveIdAsync(cancellationToken);
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
        var resolvedDriveId = driveId ?? await ResolveOwnDriveIdAsync(cancellationToken);

        return await ResolvePathItem(resolvedDriveId, path).GetAsync(cancellationToken: cancellationToken);
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
        var resolvedDriveId = driveId ?? await ResolveOwnDriveIdAsync(cancellationToken);
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
        var resolvedDriveId = driveId ?? await ResolveOwnDriveIdAsync(cancellationToken);

        await ResolvePathItem(resolvedDriveId, path).DeleteAsync(cancellationToken: cancellationToken);
    }

    private async Task<string> ResolveOwnDriveIdAsync(CancellationToken cancellationToken)
    {
        var drive = await client.Me.Drive.GetAsync(cancellationToken: cancellationToken);

        return drive?.Id ?? throw new InvalidOperationException("Could not resolve the signed-in user's OneDrive id.");
    }

    private DriveItemItemRequestBuilder ResolveRootItem(string driveId, string? path) =>
        string.IsNullOrEmpty(path)
            ? client.Drives[driveId].Items["root"]
            : client.Drives[driveId].Items["root"].ItemWithPath(path);

    private CustomDriveItemItemRequestBuilder ResolvePathItem(string driveId, string path) =>
        client.Drives[driveId].Items["root"].ItemWithPath(path);
}

/// <summary>
/// Surfaced when a folder-segment create hits Graph's 409 for an existing name -
/// the "already exists, continue" signal mkdir -p semantics (PRD §4.4) depend on.
/// </summary>
public sealed class DriveItemAlreadyExistsException(string itemName, ODataError inner)
    : Exception($"An item named '{itemName}' already exists at this location.", inner)
{
    public string ItemName { get; } = itemName;
}

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
