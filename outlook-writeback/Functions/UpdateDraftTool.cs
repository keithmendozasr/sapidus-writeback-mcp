using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Functions;

public sealed class UpdateDraftTool(OutlookGraphClient client)
{
    [Function(nameof(UpdateDraftTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger("update_draft", "Edit fields on an existing Outlook draft. Does not send it.")]
            ToolInvocationContext context,
        [McpToolProperty("draftId", "The Outlook draft's message ID.", isRequired: true)] string draftId,
        [McpToolProperty("toAddress", "New recipient email address, if changing it.")] string? toAddress,
        [McpToolProperty("subject", "New subject, if changing it.")] string? subject,
        [McpToolProperty("body", "New plain-text body, if changing it.")] string? body)
    {
        var updatedId = await client.UpdateDraftAsync(draftId, toAddress, subject, body);

        return $"Draft updated. ID: {updatedId}.";
    }
}
