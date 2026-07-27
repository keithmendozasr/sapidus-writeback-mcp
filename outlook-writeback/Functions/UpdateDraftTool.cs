using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Functions;

public sealed class UpdateDraftTool(OutlookGraphClient client)
{
    [Function(nameof(UpdateDraftTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "update_draft",
            "Edit fields on an existing Outlook draft. Does not send it. " +
                "to/cc/bcc entries must be non-blank, and the same address can't appear in more than one of the three lists - " +
                "a generic failure with no specific reason usually means one of those two checks failed, so re-check the " +
                "recipient lists for a blank or duplicated entry before retrying. Duplicate-recipient rejection only checks " +
                "the to/cc/bcc lists supplied in this call - it can't see the draft's existing recipients on fields you " +
                "leave unchanged, so an address already present in an omitted list won't be caught. " +
                "If this call fails with an authentication/401-style error, tell the user the outlook-writeback " +
                "connector may need to be reconnected (Settings/Customize > Connectors > outlook-writeback > " +
                "Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("draftId", "The Outlook draft's message ID.", isRequired: true)] string draftId,
        [McpToolProperty("to", "New recipient email address(es), if changing them. Omit to leave unchanged; pass an empty array to clear.")]
            string[]? to,
        [McpToolProperty("subject", "New subject, if changing it.")] string? subject,
        [McpToolProperty("body", "New body, if changing it. Plain text unless isHtml is true.")] string? body,
        [McpToolProperty("isHtml", "Set true if body is HTML markup instead of plain text (e.g. to include a table). Only applies when body is also provided. Defaults to false.")]
            bool? isHtml,
        [McpToolProperty("cc", "New Cc email address(es), if changing them. Omit to leave unchanged; pass an empty array to clear.")]
            string[]? cc,
        [McpToolProperty("bcc", "New Bcc email address(es), if changing them. Omit to leave unchanged; pass an empty array to clear.")]
            string[]? bcc)
    {
        var toAddresses = RecipientList.Normalize(to);
        var ccAddresses = RecipientList.Normalize(cc);
        var bccAddresses = RecipientList.Normalize(bcc);

        var duplicates = RecipientList.FindDuplicates(toAddresses ?? [], ccAddresses ?? [], bccAddresses ?? []);

        if (duplicates.Count > 0)
            throw new ArgumentException(
                $"The following address(es) appear in more than one of to/cc/bcc: {string.Join(", ", duplicates)}. Remove each from all but one list.");

        var updatedId = await client.UpdateDraftAsync(draftId, toAddresses, subject, body, isHtml ?? false, ccAddresses, bccAddresses);

        return $"Draft updated. ID: {updatedId}.";
    }
}
