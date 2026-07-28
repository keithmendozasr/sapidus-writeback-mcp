using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Functions;

public sealed class CreateDraftTool(OutlookGraphClient client)
{
    [Function(nameof(CreateDraftTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "create_draft",
            "Create an Outlook draft in the user's mailbox. Does not send it. " +
                "to/cc/bcc entries must be non-blank, and the same address can't appear in more than one of the three lists - " +
                "a generic failure with no specific reason usually means one of those two checks failed, so re-check the " +
                "recipient lists for a blank or duplicated entry before retrying. " +
                "If this call fails with an authentication/401-style error, tell the user the outlook-writeback " +
                "connector may need to be reconnected (Settings/Customize > Connectors > outlook-writeback > " +
                "Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("to", "Recipient email address(es), if any. A draft may be created with no recipients across to/cc/bcc.")]
            string[]? to,
        [McpToolProperty("subject", "Email subject.", isRequired: true)] string subject,
        [McpToolProperty("body", "Email body. Plain text unless isHtml is true.", isRequired: true)] string body,
        [McpToolProperty("isHtml", "Set true if body is HTML markup instead of plain text (e.g. to include a table). Defaults to false.")]
            bool? isHtml,
        [McpToolProperty("cc", "Cc email address(es), if any.")] string[]? cc,
        [McpToolProperty("bcc", "Bcc email address(es), if any.")] string[]? bcc)
    {
        var toAddresses = RecipientList.Normalize(to);
        var ccAddresses = RecipientList.Normalize(cc);
        var bccAddresses = RecipientList.Normalize(bcc);

        var duplicates = RecipientList.FindDuplicates(toAddresses ?? [], ccAddresses ?? [], bccAddresses ?? []);

        if (duplicates.Count > 0)
            throw new ArgumentException(
                $"The following address(es) appear in more than one of to/cc/bcc: {string.Join(", ", duplicates)}. Remove each from all but one list.");

        var draftId = await client.CreateDraftAsync(toAddresses, subject, body, isHtml ?? false, ccAddresses, bccAddresses);

        return $"Draft created. ID: {draftId}. Open it in Outlook to review and send.";
    }
}
