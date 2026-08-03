using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using DriveWriteback.Graph;
using DriveWriteback.Graph.Writes;

namespace DriveWriteback.Functions;

public sealed class CreateFileTool(DriveWriteService writeService)
{
    [Function(nameof(CreateFileTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "create_file",
            "Create a new file at a path with inline UTF-8 text content. Parents are NOT auto-created - a " +
                "typo'd path fails loudly rather than silently materializing a folder tree; call create_folder " +
                "first if the parent doesn't exist yet. conflict_behavior defaults to 'fail'. This server may be " +
                "running in dry-run mode, in which case the response is prefixed '[DRY RUN]' and nothing is " +
                "actually written.")]
            ToolInvocationContext context,
        [McpToolProperty("path", "Drive-relative path, e.g. \"Shared Documents/notes/2026-07.md\".", isRequired: true)]
            string path,
        [McpToolProperty("content", "UTF-8 text content.", isRequired: true)] string content,
        [McpToolProperty(
            "conflict_behavior",
            "'fail' (default) errors if the target already exists; 'rename' appends a numeric suffix to avoid a " +
                "collision; 'replace' overwrites the existing target.")]
            string? conflictBehavior,
        [McpToolProperty("drive_id", "Target drive ID. Omit to use the signed-in user's own OneDrive.")]
            string? driveId)
    {
        var resolvedConflictBehavior = conflictBehavior ?? "fail";
        var result = await writeService.CreateFileAsync(path, content, resolvedConflictBehavior, driveId);

        if (!result.DryRun)
        {
            var (_, requestedName) = DrivePath.SplitParent(path);
            var renameNote = resolvedConflictBehavior == "rename" && result.Item?.Name is { } actualName && actualName != requestedName
                ? $" Renamed to \"{actualName}\" to avoid a collision."
                : "";

            return $"File \"{path}\" created.{renameNote} ID: {result.Item?.Id}. Size: {result.Item?.Size} bytes. " +
                $"ETag: {result.Item?.ETag}. Web URL: {result.Item?.WebUrl}.";
        }

        if (result.TargetAlreadyExisted)
            return $"[DRY RUN] \"{path}\" already exists; conflict_behavior={resolvedConflictBehavior} would apply. Not written.";

        return $"[DRY RUN] Parent confirmed to exist, would create \"{path}\". Not written.";
    }
}
