using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Functions;

public sealed class CreateDraftTool(OutlookGraphClient client)
{
    [Function(nameof(CreateDraftTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger("create_draft", "Create a plain-text Outlook draft in the user's mailbox. Does not send it.")]
            ToolInvocationContext context,
        [McpToolProperty("toAddress", "Recipient email address.", isRequired: true)] string toAddress,
        [McpToolProperty("subject", "Email subject.", isRequired: true)] string subject,
        [McpToolProperty("body", "Plain-text email body.", isRequired: true)] string body)
    {
        var draftId = await client.CreateDraftAsync(toAddress, subject, body);

        return $"Draft created. ID: {draftId}. Open it in Outlook to review and send.";
    }
}
