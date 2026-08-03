using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item;
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
