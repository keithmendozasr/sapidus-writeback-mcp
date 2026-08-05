using DriveWriteback.Graph.Writes;
using Microsoft.Graph.Models;
using Sapidus.Writeback.Shared.Confirmation;

namespace DriveWriteback.Graph.Confirmation;

/// <summary>
/// Orchestrates delete_item's two-step, confirmation-gated flow (PRD §4.4, §9 item 3), mirroring
/// outlook-writeback's delete_event posture: the first call resolves the target, checks
/// expected_name/recursive, and issues a confirmation token without deleting anything; the
/// second call only deletes once that same token validates for the same item id.
///
/// Reads go straight to DriveGraphClient, same as every other read in this codebase. The actual
/// delete routes through DriveWriteService.DeleteItemAsync instead, so it inherits dry-run
/// gating the same way every other write does - unlike delete_event, which has no dry-run
/// concept to thread through.
///
/// expected_name and the recursive/ChildCount guard are both re-checked at confirm time, not
/// just at preview time: the token only binds an item id, so a rename or a folder gaining
/// children in the window between the two calls must fail closed rather than being silently
/// trusted from the preview.
/// </summary>
public sealed class DriveItemDeletionService(
    DriveGraphClient client,
    DriveWriteService writeService,
    ConfirmationTokenService tokenService)
{
    public async Task<PendingItemDeletion> RequestDeletionAsync(
        string pathOrId,
        string expectedName,
        bool recursive,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var (item, resolvedAsId) = await client.GetItemAsync(pathOrId, driveId, cancellationToken);

        if (item is null)
            throw new ItemNotFoundException(pathOrId);

        EnsureGuardsPass(item, expectedName, recursive);

        var token = tokenService.Issue(item.Id!);

        return new PendingItemDeletion(item.Id!, item.Name!, resolvedAsId, token);
    }

    public async Task<ConfirmedItemDeletion> ConfirmDeletionAsync(
        string itemId,
        string confirmationToken,
        string expectedName,
        bool recursive,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        if (!tokenService.Validate(itemId, confirmationToken))
            return ConfirmedItemDeletion.InvalidToken;

        var item = await client.GetItemByIdAsync(itemId, driveId, cancellationToken);

        if (item is null)
            return ConfirmedItemDeletion.Deleted; // PRD §7: already-deleted is success.

        EnsureGuardsPass(item, expectedName, recursive);

        var result = await writeService.DeleteItemAsync(itemId, driveId, cancellationToken);

        return result.DryRun ? ConfirmedItemDeletion.DryRun : ConfirmedItemDeletion.Deleted;
    }

    private static void EnsureGuardsPass(DriveItem item, string expectedName, bool recursive)
    {
        if (item.Name != expectedName)
            throw new DriveItemNameMismatchException(expectedName, item.Name ?? "");

        if (item.Folder is not null && item.Folder.ChildCount is > 0 && !recursive)
            throw new DriveFolderNotEmptyException(item.Name ?? "", item.Folder.ChildCount!.Value);
    }
}

/// <summary>
/// Result of DriveItemDeletionService.RequestDeletionAsync - nothing has been deleted yet.
/// ResolvedAsId records whether pathOrId was interpreted as an id or a path (same discriminator
/// get_item reports), so the tool response can tell the caller which interpretation was used.
/// </summary>
public sealed record PendingItemDeletion(string ItemId, string Name, bool ResolvedAsId, string ConfirmationToken);

/// <summary>
/// Result of DriveItemDeletionService.ConfirmDeletionAsync. DryRun means the token validated
/// and both guards passed, but the actual delete was skipped because the server is running in
/// dry-run mode - distinct from Deleted so the tool response can say so plainly.
/// </summary>
public enum ConfirmedItemDeletion
{
    Deleted,
    DryRun,
    InvalidToken,
}

/// <summary>
/// Surfaced when expected_name doesn't match the resolved item's actual name - the drift guard
/// PRD §4.4 requires, re-checked at both the preview and confirm calls. Meaningful only to the
/// delete confirmation flow, not general Graph-layer vocabulary, so it lives here rather than
/// alongside DriveGraphClient's exceptions.
/// </summary>
public sealed class DriveItemNameMismatchException(string expectedName, string actualName)
    : Exception($"Expected name '{expectedName}' does not match the resolved item's actual name '{actualName}'.")
{
    public string ExpectedName { get; } = expectedName;
    public string ActualName { get; } = actualName;
}
