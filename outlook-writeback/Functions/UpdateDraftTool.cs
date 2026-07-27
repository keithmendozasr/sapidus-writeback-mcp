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
                "If this call fails with an authentication/401-style error, tell the user the outlook-writeback " +
                "connector may need to be reconnected (Settings/Customize > Connectors > outlook-writeback > " +
                "Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("draftId", "The Outlook draft's message ID.", isRequired: true)] string draftId,
        [McpToolProperty("toAddress", "New recipient email address, if changing it.")] string? toAddress,
        [McpToolProperty("subject", "New subject, if changing it.")] string? subject,
        [McpToolProperty("body", "New body, if changing it. Plain text unless isHtml is true.")] string? body,
        [McpToolProperty("isHtml", "Set true if body is HTML markup instead of plain text (e.g. to include a table). Only applies when body is also provided. Defaults to false.")]
            bool? isHtml)
    {
        var updatedId = await client.UpdateDraftAsync(draftId, toAddress, subject, body, isHtml ?? false);

        return $"Draft updated. ID: {updatedId}.";
    }
}
