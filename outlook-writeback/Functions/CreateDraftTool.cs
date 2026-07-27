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
                "If this call fails with an authentication/401-style error, tell the user the outlook-writeback " +
                "connector may need to be reconnected (Settings/Customize > Connectors > outlook-writeback > " +
                "Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("toAddress", "Recipient email address.", isRequired: true)] string toAddress,
        [McpToolProperty("subject", "Email subject.", isRequired: true)] string subject,
        [McpToolProperty("body", "Email body. Plain text unless isHtml is true.", isRequired: true)] string body,
        [McpToolProperty("isHtml", "Set true if body is HTML markup instead of plain text (e.g. to include a table). Defaults to false.")]
            bool? isHtml)
    {
        var draftId = await client.CreateDraftAsync(toAddress, subject, body, isHtml ?? false);

        return $"Draft created. ID: {draftId}. Open it in Outlook to review and send.";
    }
}
