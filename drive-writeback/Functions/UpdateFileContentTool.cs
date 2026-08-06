using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using DriveWriteback.Graph.Writes;

namespace DriveWriteback.Functions;

public sealed class UpdateFileContentTool(DriveWriteService writeService)
{
    [Function(nameof(UpdateFileContentTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "update_file_content",
            "Replace the full content of an existing file. if_match is REQUIRED - the eTag or cTag from a prior " +
                "get_item call - and enforced server-side by Graph as an optimistic-concurrency check: if the file " +
                "has changed since that read, the call fails with a distinct error rather than silently overwriting " +
                "someone else's change. On that error, call get_item again for the current eTag and retry. This " +
                "server may be running in dry-run mode, in which case the response is prefixed '[DRY RUN]' and " +
                "nothing is actually written.")]
            ToolInvocationContext context,
        [McpToolProperty("path_or_id", "Drive-relative path or a Graph item ID. Both are accepted anywhere the other is.", isRequired: true)]
            string pathOrId,
        [McpToolProperty("content", "UTF-8 text content to replace the file's current content with.", isRequired: true)]
            string content,
        [McpToolProperty("if_match", "The eTag or cTag from a prior get_item call. Required.", isRequired: true)]
            string ifMatch,
        [McpToolProperty("drive_id", "Target drive ID. Omit to use the signed-in user's own OneDrive. Another user's OneDrive (e.g. an item they've shared with you) is allowed; SharePoint document library drives are rejected - not supported by this deployment.")]
            string? driveId)
    {
        var result = await writeService.UpdateFileContentAsync(pathOrId, content, ifMatch, driveId);

        if (result.DryRun)
            return $"[DRY RUN] Would update \"{pathOrId}\" (item ID {result.Item?.Id}) using if_match={ifMatch}. Not written.";

        return $"File \"{pathOrId}\" updated. ID: {result.Item?.Id}. New size: {result.Item?.Size} bytes. " +
            $"New ETag: {result.Item?.ETag}.";
    }
}
