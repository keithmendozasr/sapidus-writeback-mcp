using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using DriveWriteback.Graph.Writes;

namespace DriveWriteback.Functions;

public sealed class CreateFolderTool(DriveWriteService writeService)
{
    [Function(nameof(CreateFolderTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "create_folder",
            "Create a folder, mkdir -p semantics: walks the path from the drive root, creating each missing " +
                "segment. Idempotent - an already-existing path succeeds with no changes rather than an error. " +
                "This server may be running in dry-run mode, in which case the response is prefixed '[DRY RUN]' " +
                "and nothing is actually created.")]
            ToolInvocationContext context,
        [McpToolProperty("path", "Full drive-relative path of the folder to create, e.g. \"Shared Documents/notes/2026\".", isRequired: true)]
            string path,
        [McpToolProperty("drive_id", "Target drive ID. Omit to use the signed-in user's own OneDrive.")]
            string? driveId)
    {
        var result = await writeService.CreateFolderAsync(path, driveId);

        if (!result.DryRun)
            return $"Folder \"{path}\" ready. ID: {result.Item?.Id}. ETag: {result.Item?.ETag}.";

        if (result.AlreadyExisted)
            return $"[DRY RUN] \"{path}\" already exists, no changes made.";

        return $"[DRY RUN] Would create {result.CreatedSegments.Count} new folder segment(s) under \"{path}\": " +
            $"{string.Join(", ", result.CreatedSegments)}. Not created - dry-run mode is on.";
    }
}
