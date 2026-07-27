using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Functions;

public sealed class UpdateEventTool(OutlookGraphClient client)
{
    [Function(nameof(UpdateEventTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "update_event",
            "Edit fields on an existing calendar event. attendees entries must be non-blank - a generic failure with no " +
                "specific reason usually means a blank entry slipped into attendees, so re-check it before retrying. " +
                "If this call fails with an authentication/401-style error, tell the user the outlook-writeback " +
                "connector may need to be reconnected (Settings/Customize > Connectors > outlook-writeback > " +
                "Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("eventId", "The calendar event's ID.", isRequired: true)] string eventId,
        [McpToolProperty("subject", "New subject/title, if changing it.")] string? subject,
        [McpToolProperty("start", "New start time, ISO 8601 with a timezone offset, if changing it.")] string? start,
        [McpToolProperty("end", "New end time, ISO 8601 with a timezone offset, if changing it.")] string? end,
        [McpToolProperty("location", "New location, if changing it.")] string? location,
        [McpToolProperty("body", "New notes/description, if changing it.")] string? body,
        [McpToolProperty(
            "attendees",
            "New attendee email addresses, if changing them. Omit to leave the attendee list unchanged; pass an empty array to clear it entirely. " +
                "Note: this is a JSON array, unlike create_event's attendees which is a comma-separated string - the two tools intentionally use different wire types for the same field name.")]
            string[]? attendees)
    {
        var attendeeAddresses = RecipientList.Normalize(attendees);

        var updatedId = await client.UpdateEventAsync(
            eventId,
            subject,
            start is null ? null : DateTimeOffset.Parse(start),
            end is null ? null : DateTimeOffset.Parse(end),
            location,
            body,
            attendeeAddresses);

        return $"Event updated. ID: {updatedId}.";
    }
}
