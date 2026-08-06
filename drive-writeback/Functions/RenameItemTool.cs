using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using DriveWriteback.Graph.Writes;

namespace DriveWriteback.Functions;

public sealed class RenameItemTool(DriveWriteService writeService)
{
    [Function(nameof(RenameItemTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "rename_item",
            "Change a file or folder's name in place, leaving it in the same parent folder. This server may be " +
                "running in dry-run mode, in which case the response is prefixed '[DRY RUN]' and nothing is " +
                "actually renamed.")]
            ToolInvocationContext context,
        [McpToolProperty("path_or_id", "Drive-relative path or a Graph item ID. Both are accepted anywhere the other is.", isRequired: true)]
            string pathOrId,
        [McpToolProperty("new_name", "The item's new name.", isRequired: true)] string newName,
        [McpToolProperty("drive_id", "Target drive ID. Omit to use the signed-in user's own OneDrive. Another user's OneDrive (e.g. an item they've shared with you) is allowed; SharePoint document library drives are rejected - not supported by this deployment.")]
            string? driveId)
    {
        var result = await writeService.RenameItemAsync(pathOrId, newName, driveId);

        if (result.DryRun)
            return $"[DRY RUN] Would rename \"{pathOrId}\" (item ID {result.Item?.Id}) to \"{newName}\". Not renamed.";

        return $"\"{pathOrId}\" renamed to \"{newName}\". ID: {result.Item?.Id}.";
    }
}
