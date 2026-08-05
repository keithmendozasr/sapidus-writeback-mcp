using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using DriveWriteback.Graph;

namespace DriveWriteback.Functions;

public sealed class GetItemTool(DriveGraphClient client)
{
    [Function(nameof(GetItemTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "get_item",
            "Fetch metadata for one file or folder: eTag, cTag, size, folder/file discriminator, web URL. " +
                "This is a pre-write safety check, not a browsing tool - it does not enumerate a folder's contents. " +
                "A target that doesn't exist surfaces as an error, not a negative answer - parents are never " +
                "auto-created by create_file, so checking here first before create_folder/create_file avoids a " +
                "failed write on a typo'd path.")]
            ToolInvocationContext context,
        [McpToolProperty(
            "path_or_id",
            "Drive-relative path (e.g. \"notes/2026-07.md\") or a Graph item ID. Both are accepted anywhere the other is.",
            isRequired: true)]
            string pathOrId,
        [McpToolProperty("drive_id", "Target drive ID. Omit to use the signed-in user's own OneDrive. Another user's OneDrive (e.g. an item they've shared with you) is allowed; SharePoint document library drives are rejected - not supported by this deployment.")]
            string? driveId)
    {
        var (item, resolvedAsId) = await client.GetItemAsync(pathOrId, driveId);
        var kind = item!.Folder is not null ? "folder" : "file";
        var interpretation = resolvedAsId ? "ID" : "path";

        return $"{kind} \"{item.Name}\" (resolved \"{pathOrId}\" as {interpretation}). " +
            $"ID: {item.Id}. Size: {item.Size} bytes. ETag: {item.ETag}. CTag: {item.CTag}. Web URL: {item.WebUrl}. " +
            "Checkout state: not tracked for OneDrive in this phase.";
    }
}
