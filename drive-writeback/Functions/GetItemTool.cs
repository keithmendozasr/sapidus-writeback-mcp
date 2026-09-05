using System.Globalization;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using Microsoft.Extensions.Logging;
using DriveWriteback.Graph;

namespace DriveWriteback.Functions;

public sealed class GetItemTool(DriveGraphClient client, ILogger<GetItemTool> logger)
{
    [Function(nameof(GetItemTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "get_item",
            "Fetch metadata for one file or folder: size, folder/file discriminator, web URL, and an if_match value " +
                "for update_file_content - copy that value verbatim, including its surrounding double quotes, into " +
                "if_match; dropping the quotes makes update_file_content fail every time regardless of how fresh " +
                "the read was. Scoped to edit workflows, not general browsing - it does not enumerate a folder's " +
                "contents, and requires not_modified_since, a timestamp anchored to a prior content read (see that " +
                "parameter's own description). A target that doesn't exist surfaces as an error, not a negative " +
                "answer - parents are never auto-created by create_file, so checking here first before " +
                "create_folder/create_file avoids a failed write on a typo'd path.")]
            ToolInvocationContext context,
        [McpToolProperty(
            "path_or_id",
            "Drive-relative path (e.g. \"notes/2026-07.md\") or a Graph item ID. Both are accepted anywhere the other is.",
            isRequired: true)]
            string pathOrId,
        [McpToolProperty(
            "not_modified_since",
            "ISO 8601 timestamp marking the moment you read this file's content - REQUIRED. Prefer the file's own " +
                "last-modified time if your read tool surfaced one (exact match against this server's clock, no " +
                "skew). Otherwise use your own wall-clock time at the moment you read the content - weaker, since a " +
                "few seconds of drift can mask a real conflict or cause a spurious rejection. Call get_item " +
                "concurrently with that content read, never right before writing - a same-call get_item always " +
                "matches itself and provides no protection. If the file was modified after this timestamp, this " +
                "call fails instead of returning an if_match value; re-read the file's content and start over.",
            isRequired: true)]
            string notModifiedSinceRaw,
        [McpToolProperty("drive_id", "Target drive ID. Omit to use the signed-in user's own OneDrive. Another user's OneDrive (e.g. an item they've shared with you) is allowed; SharePoint document library drives are rejected - not supported by this deployment.")]
            string? driveId)
    {
        if (!DateTimeOffset.TryParse(notModifiedSinceRaw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var notModifiedSince))
            throw new ArgumentException($"not_modified_since \"{notModifiedSinceRaw}\" is not a valid ISO 8601 timestamp.");

        try
        {
            var (item, resolvedAsId) = await client.GetItemAsync(pathOrId, driveId, notModifiedSince);

            logger.LogInformation(
                "get_item {PathOrId}: outcome=pass drive-id={DriveId} not-modified-since={NotModifiedSince} last-modified={LastModifiedDateTime}",
                pathOrId, driveId, notModifiedSince, item!.LastModifiedDateTime);

            var kind = item.Folder is not null ? "folder" : "file";
            var interpretation = resolvedAsId ? "ID" : "path";

            return $"{kind} \"{item.Name}\" (resolved \"{pathOrId}\" as {interpretation}). " +
                $"ID: {item.Id}. Size: {item.Size} bytes. Web URL: {item.WebUrl}. " +
                $"if_match value for update_file_content - copy this exact string verbatim into that parameter, " +
                $"including the double-quote characters at each end (they are part of the value Graph checks, not " +
                $"sentence punctuation): {item.ETag} " +
                $"-- the item's cTag is an equally valid if_match value, same copy-verbatim rule: {item.CTag} " +
                "Checkout state: not tracked for OneDrive in this phase.";
        }
        catch (ItemModifiedSinceReadException ex)
        {
            logger.LogInformation(
                "get_item {PathOrId}: outcome=rejected drive-id={DriveId} not-modified-since={NotModifiedSince} last-modified={LastModifiedDateTime}",
                ex.PathOrId, driveId, ex.NotModifiedSince, ex.LastModifiedDateTime);

            throw;
        }
    }
}
