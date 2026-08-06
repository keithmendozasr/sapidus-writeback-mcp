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
            "Fetch metadata for one file or folder: size, folder/file discriminator, web URL, and an if_match value " +
                "for update_file_content - copy that value verbatim, including its surrounding double quotes, into " +
                "if_match; dropping the quotes makes update_file_content fail every time regardless of how fresh " +
                "the read was. This is a pre-write safety check, not a browsing tool - it does not enumerate a " +
                "folder's contents. A target that doesn't exist surfaces as an error, not a negative answer - " +
                "parents are never auto-created by create_file, so checking here first before create_folder/" +
                "create_file avoids a failed write on a typo'd path.")]
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
            $"ID: {item.Id}. Size: {item.Size} bytes. Web URL: {item.WebUrl}. " +
            $"if_match value for update_file_content - copy this exact string verbatim into that parameter, " +
            $"including the double-quote characters at each end (they are part of the value Graph checks, not " +
            $"sentence punctuation): {item.ETag} " +
            $"-- the item's cTag is an equally valid if_match value, same copy-verbatim rule: {item.CTag} " +
            "Checkout state: not tracked for OneDrive in this phase.";
    }
}
