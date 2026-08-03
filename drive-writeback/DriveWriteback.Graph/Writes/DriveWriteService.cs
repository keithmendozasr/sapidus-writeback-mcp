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
