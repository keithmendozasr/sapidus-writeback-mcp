using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;

namespace DriveWriteback.Graph.Writes;

/// <summary>
/// The only place dry-run logic lives - DriveGraphClient stays a dumb, dry-run-agnostic
/// Graph wrapper. Every call path (dry-run and real) ends in one structured audit log
/// entry (PRD §9 item 7's minimal version - no caller-OID field yet, since there's no
/// Easy Auth token to pull it from until boundary-7b exists).
/// </summary>
public sealed class DriveWriteService(DriveGraphClient client, DriveWriteOptions options, ILogger<DriveWriteService> logger)
{
    /// <summary>
    /// mkdir -p. In dry-run mode, resolves as far as existing folders go and reports what
    /// would be created without ever POSTing (DriveGraphClient.ResolveFolderPathAsync).
    /// </summary>
    public async Task<CreateFolderResult> CreateFolderAsync(
        string path,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        DrivePath.Validate(path);

        var resolution = await client.ResolveFolderPathAsync(path, driveId, cancellationToken);
        var alreadyExisted = resolution.MissingSegments.Count == 0;

        if (options.DryRun)
        {
            logger.LogInformation(
                "create_folder {Path}: outcome=dry-run already-exists={AlreadyExisted} missing-segments={MissingSegments}",
                path, alreadyExisted, string.Join(',', resolution.MissingSegments));

            return new CreateFolderResult(DryRun: true, alreadyExisted, resolution.MissingSegments, resolution.DeepestExisting);
        }

        var created = await client.CreateFolderPathAsync(path, driveId, cancellationToken);

        logger.LogInformation(
            "create_folder {Path}: outcome=created item-id={ItemId} etag={ETag}",
            path, created?.Id, created?.ETag);

        return new CreateFolderResult(DryRun: false, alreadyExisted, resolution.MissingSegments, created);
    }

    /// <summary>
    /// Validates content size and calls DrivePath.Validate before checking options.DryRun -
    /// those are validation, not mutation, and must fire in dry-run too. In dry-run mode,
    /// re-implements CreateFileAsync's own parent-exists/target-exists checks read-only
    /// (via TryGetItemByPathAsync) rather than calling the mutating method; in real mode,
    /// delegates entirely to DriveGraphClient.CreateFileAsync and lets its own exceptions
    /// surface.
    /// </summary>
    public async Task<CreateFileResult> CreateFileAsync(
        string path,
        string content,
        string conflictBehavior = "fail",
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        DrivePath.Validate(path);

        var contentBytes = Encoding.UTF8.GetByteCount(content);

        if (contentBytes > options.MaxContentBytes)
            throw new DriveContentTooLargeException(contentBytes, options.MaxContentBytes);

        if (!options.DryRun)
        {
            var created = await client.CreateFileAsync(path, content, conflictBehavior, driveId, cancellationToken);

            logger.LogInformation(
                "create_file {Path}: outcome=created conflict-behavior={ConflictBehavior} item-id={ItemId} etag={ETag}",
                path, conflictBehavior, created?.Id, created?.ETag);

            return new CreateFileResult(DryRun: false, TargetAlreadyExisted: false, conflictBehavior, created);
        }

        var (parentPath, name) = DrivePath.SplitParent(path);

        if (parentPath.Length > 0 && await client.TryGetItemByPathAsync(parentPath, driveId, cancellationToken) is null)
            throw new DriveParentNotFoundException(parentPath);

        var existingTarget = await client.TryGetItemByPathAsync(path, driveId, cancellationToken);

        if (existingTarget is not null && conflictBehavior == "fail")
            throw new DriveItemAlreadyExistsException(name);

        logger.LogInformation(
            "create_file {Path}: outcome=dry-run conflict-behavior={ConflictBehavior} target-exists={TargetExists}",
            path, conflictBehavior, existingTarget is not null);

        return new CreateFileResult(DryRun: true, existingTarget is not null, conflictBehavior, Item: null);
    }

    /// <summary>
    /// Replaces an existing file's content. if_match is mandatory (PRD §4.4) - whole-file
    /// replacement without optimistic concurrency is a silent-data-loss machine. Resolves
    /// path_or_id via GetItemAsync first (id-preferred addressing, same reasoning as every
    /// other Phase 2 method) and acts by id. if_match doubles as the eTag-before value for
    /// the audit log - folding it into the one log line here closes that part of PRD §9 item
    /// 7's gap without inventing a second logging shape.
    /// </summary>
    public async Task<UpdateFileContentResult> UpdateFileContentAsync(
        string pathOrId,
        string content,
        string ifMatch,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var contentBytes = Encoding.UTF8.GetByteCount(content);

        if (contentBytes > options.MaxContentBytes)
            throw new DriveContentTooLargeException(contentBytes, options.MaxContentBytes);

        var (item, _) = await client.GetItemAsync(pathOrId, driveId, cancellationToken: cancellationToken);

        if (item is null)
            throw new ItemNotFoundException(pathOrId);

        if (options.DryRun)
        {
            logger.LogInformation(
                "update_file_content {PathOrId}: outcome=dry-run item-id={ItemId} if-match={IfMatch}",
                pathOrId, item.Id, ifMatch);

            return new UpdateFileContentResult(DryRun: true, item);
        }

        var updated = await client.ReplaceContentByIdAsync(item.Id!, content, ifMatch, driveId, cancellationToken);

        logger.LogInformation(
            "update_file_content {PathOrId}: outcome=updated item-id={ItemId} if-match={IfMatch} etag-after={ETag}",
            pathOrId, item.Id, ifMatch, updated?.ETag);

        return new UpdateFileContentResult(DryRun: false, updated);
    }

    /// <summary>
    /// Renames an item in place. Resolves path_or_id via GetItemAsync first, acts by id.
    /// </summary>
    public async Task<RenameItemResult> RenameItemAsync(
        string pathOrId,
        string newName,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        DrivePath.Validate(newName);

        var (item, _) = await client.GetItemAsync(pathOrId, driveId, cancellationToken: cancellationToken);

        if (item is null)
            throw new ItemNotFoundException(pathOrId);

        if (options.DryRun)
        {
            logger.LogInformation(
                "rename_item {PathOrId}: outcome=dry-run item-id={ItemId} new-name={NewName}",
                pathOrId, item.Id, newName);

            return new RenameItemResult(DryRun: true, item, newName);
        }

        var renamed = await client.RenameItemAsync(item.Id!, newName, driveId, cancellationToken);

        logger.LogInformation(
            "rename_item {PathOrId}: outcome=renamed item-id={ItemId} new-name={NewName}",
            pathOrId, item.Id, newName);

        return new RenameItemResult(DryRun: false, renamed, newName);
    }

    /// <summary>
    /// Moves an item to a new parent folder, same drive only (PRD §4.4). Resolves both the
    /// item and the destination parent via GetItemAsync first, acts by id.
    /// </summary>
    public async Task<MoveItemResult> MoveItemAsync(
        string pathOrId,
        string newParentPathOrId,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        var (item, _) = await client.GetItemAsync(pathOrId, driveId, cancellationToken: cancellationToken);

        if (item is null)
            throw new ItemNotFoundException(pathOrId);

        var (parent, _) = await client.GetItemAsync(newParentPathOrId, driveId, cancellationToken: cancellationToken);

        if (parent is null)
            throw new DriveParentNotFoundException(newParentPathOrId);

        if (options.DryRun)
        {
            logger.LogInformation(
                "move_item {PathOrId}: outcome=dry-run item-id={ItemId} new-parent-id={ParentId}",
                pathOrId, item.Id, parent.Id);

            return new MoveItemResult(DryRun: true, item, parent.Id!);
        }

        var moved = await client.MoveItemAsync(item.Id!, parent.Id!, driveId, cancellationToken);

        logger.LogInformation(
            "move_item {PathOrId}: outcome=moved item-id={ItemId} new-parent-id={ParentId}",
            pathOrId, item.Id, parent.Id);

        return new MoveItemResult(DryRun: false, moved, parent.Id!);
    }

    /// <summary>
    /// The delete primitive DriveItemDeletionService's confirmed second call delegates to,
    /// once its own expected_name/recursive guards have already passed. Internal, not a
    /// directly-exposed tool method - mirrors delete_item's confirmation-gated posture, where
    /// there's no single-call "just delete this" surface. Takes an id, not path_or_id -
    /// DriveItemDeletionService has already resolved the target by the time this is called.
    /// </summary>
    internal async Task<DeleteItemResult> DeleteItemAsync(
        string itemId,
        string? driveId = null,
        CancellationToken cancellationToken = default)
    {
        if (options.DryRun)
        {
            logger.LogInformation("delete_item: outcome=dry-run item-id={ItemId}", itemId);

            return new DeleteItemResult(DryRun: true);
        }

        await client.DeleteItemByIdAsync(itemId, driveId, cancellationToken);

        logger.LogInformation("delete_item: outcome=deleted item-id={ItemId}", itemId);

        return new DeleteItemResult(DryRun: false);
    }
}

/// <summary>
/// Result of DriveWriteService.CreateFolderAsync: AlreadyExisted is true when
/// ResolveFolderPathAsync found the full path already there (dry-run and real mode agree
/// on this value); CreatedSegments lists what was/would be created; Item is the deepest
/// existing folder in dry-run mode, or the final created folder in real mode.
/// </summary>
public sealed record CreateFolderResult(bool DryRun, bool AlreadyExisted, IReadOnlyList<string> CreatedSegments, DriveItem? Item);

/// <summary>
/// Result of DriveWriteService.CreateFileAsync: TargetAlreadyExisted reports whether a
/// pre-existing item was found at the target path (dry-run mode only - real mode always
/// reports false here since a real "fail" conflict throws instead of returning a result).
/// Item is null in dry-run mode (nothing was written) and the created item in real mode.
/// </summary>
public sealed record CreateFileResult(bool DryRun, bool TargetAlreadyExisted, string ConflictBehavior, DriveItem? Item);

/// <summary>
/// Result of DriveWriteService.UpdateFileContentAsync. Item is the pre-update item in
/// dry-run mode (nothing was written) and the post-update item in real mode.
/// </summary>
public sealed record UpdateFileContentResult(bool DryRun, DriveItem? Item);

/// <summary>
/// Result of DriveWriteService.RenameItemAsync. Item is the pre-rename item in dry-run mode
/// and the post-rename item in real mode; NewName is always the requested name, regardless
/// of mode.
/// </summary>
public sealed record RenameItemResult(bool DryRun, DriveItem? Item, string NewName);

/// <summary>
/// Result of DriveWriteService.MoveItemAsync. Item is the pre-move item in dry-run mode and
/// the post-move item in real mode; NewParentId is always the resolved destination parent's
/// id, regardless of mode.
/// </summary>
public sealed record MoveItemResult(bool DryRun, DriveItem? Item, string NewParentId);

/// <summary>
/// Result of DriveWriteService.DeleteItemAsync.
/// </summary>
public sealed record DeleteItemResult(bool DryRun);
