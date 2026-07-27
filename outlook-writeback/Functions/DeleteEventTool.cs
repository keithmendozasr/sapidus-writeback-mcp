using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using OutlookWriteback.Graph.Confirmation;

namespace OutlookWriteback.Functions;

public sealed class DeleteEventTool(EventDeletionService deletionService)
{
    [Function(nameof(DeleteEventTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "delete_event",
            "Delete a calendar event. Two-step and confirmation-gated: call once with just " +
                "eventId to preview the event and get a confirmationToken - this does NOT delete " +
                "anything. Call again with the same eventId and that confirmationToken to actually delete.")]
            ToolInvocationContext context,
        [McpToolProperty("eventId", "The calendar event's ID.", isRequired: true)] string eventId,
        [McpToolProperty("confirmationToken", "Omit on the first call. Supply the token returned by the first call to confirm the delete.")]
            string? confirmationToken)
    {
        if (confirmationToken is null)
        {
            var pending = await deletionService.RequestDeletionAsync(eventId);

            return $"About to delete \"{pending.Subject}\" (event ID: {eventId}). " +
                "This has NOT been deleted yet. If the user confirms, call delete_event again with " +
                $"eventId=\"{eventId}\" and confirmationToken=\"{pending.ConfirmationToken}\" to delete it.";
        }

        var confirmed = await deletionService.ConfirmDeletionAsync(eventId, confirmationToken);

        return confirmed
            ? $"Event {eventId} deleted."
            : "That confirmation token is invalid or expired. Call delete_event again with just eventId to get a new one.";
    }
}
