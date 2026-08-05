using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using DriveWriteback.Graph.Confirmation;

namespace DriveWriteback.Functions;

public sealed class DeleteItemTool(DriveItemDeletionService deletionService)
{
    [Function(nameof(DeleteItemTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "delete_item",
            "Delete a file or folder, moving it to the recycle bin - never a permanent delete. Two-step and " +
                "confirmation-gated: call once with just path_or_id/expected_name/recursive to preview the item " +
                "and get a confirmation_token - this does NOT delete anything. Call again with the same " +
                "expected_name/recursive plus that confirmation_token to actually delete. expected_name must match " +
                "the resolved item's actual name - it guards against path/ID drift between the read and the write. " +
                "recursive defaults to false, so deleting a non-empty folder fails rather than silently taking a " +
                "subtree - pass recursive=true to allow it. This server may be running in dry-run mode, in which " +
                "case the confirm call is still token-gated but the response is prefixed '[DRY RUN]' and nothing " +
                "is actually deleted.")]
            ToolInvocationContext context,
        [McpToolProperty("path_or_id", "Drive-relative path or a Graph item ID.", isRequired: true)]
            string pathOrId,
        [McpToolProperty("expected_name", "Must match the resolved item's actual name.", isRequired: true)]
            string expectedName,
        [McpToolProperty("recursive", "Required true to delete a non-empty folder. Defaults to false.")]
            bool? recursive,
        [McpToolProperty("confirmation_token", "Omit on the first call. Supply the token returned by the first call to confirm the delete.")]
            string? confirmationToken,
        [McpToolProperty("drive_id", "Target drive ID. Omit to use the signed-in user's own OneDrive.")]
            string? driveId)
    {
        var resolvedRecursive = recursive ?? false;

        if (confirmationToken is null)
        {
            var pending = await deletionService.RequestDeletionAsync(pathOrId, expectedName, resolvedRecursive, driveId);

            return $"About to delete \"{pending.Name}\" (item ID: {pending.ItemId}). This has NOT been deleted yet. " +
                "If the user confirms this is the right item, call delete_item again with " +
                $"path_or_id=\"{pending.ItemId}\", expected_name=\"{pending.Name}\", recursive={resolvedRecursive.ToString().ToLowerInvariant()}, " +
                $"and confirmation_token=\"{pending.ConfirmationToken}\" to delete it.";
        }

        var confirmed = await deletionService.ConfirmDeletionAsync(pathOrId, confirmationToken, expectedName, resolvedRecursive, driveId);

        return confirmed switch
        {
            ConfirmedItemDeletion.Deleted => $"Item {pathOrId} deleted.",
            ConfirmedItemDeletion.DryRun => $"[DRY RUN] Would delete item {pathOrId}. Not deleted.",
            _ => "That confirmation token is invalid or expired. Call delete_item again with just path_or_id, " +
                "expected_name, and recursive to get a new one.",
        };
    }
}
