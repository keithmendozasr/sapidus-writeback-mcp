using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using DriveWriteback.Graph.Writes;

namespace DriveWriteback.Functions;

public sealed class MoveItemTool(DriveWriteService writeService)
{
    [Function(nameof(MoveItemTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "move_item",
            "Move a file or folder into a different parent folder, within the same drive only - cross-drive " +
                "moves are not supported and fail with a clear error. This server may be running in dry-run mode, " +
                "in which case the response is prefixed '[DRY RUN]' and nothing is actually moved.")]
            ToolInvocationContext context,
        [McpToolProperty("path_or_id", "Drive-relative path or a Graph item ID of the item to move.", isRequired: true)]
            string pathOrId,
        [McpToolProperty("new_parent_path_or_id", "Drive-relative path or a Graph item ID of the destination folder.", isRequired: true)]
            string newParentPathOrId,
        [McpToolProperty("drive_id", "Target drive ID. Omit to use the signed-in user's own OneDrive.")]
            string? driveId)
    {
        var result = await writeService.MoveItemAsync(pathOrId, newParentPathOrId, driveId);

        if (result.DryRun)
            return $"[DRY RUN] Would move \"{pathOrId}\" (item ID {result.Item?.Id}) into \"{newParentPathOrId}\" " +
                $"(parent ID {result.NewParentId}). Not moved.";

        return $"\"{pathOrId}\" moved into \"{newParentPathOrId}\" (parent ID {result.NewParentId}). " +
            $"Item ID: {result.Item?.Id}.";
    }
}
