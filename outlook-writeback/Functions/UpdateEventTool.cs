using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Functions;

public sealed class UpdateEventTool(OutlookGraphClient client)
{
    [Function(nameof(UpdateEventTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger("update_event", "Edit fields on an existing calendar event.")]
            ToolInvocationContext context,
        [McpToolProperty("eventId", "The calendar event's ID.", isRequired: true)] string eventId,
        [McpToolProperty("subject", "New subject/title, if changing it.")] string? subject,
        [McpToolProperty("start", "New start time, ISO 8601 with a timezone offset, if changing it.")] string? start,
        [McpToolProperty("end", "New end time, ISO 8601 with a timezone offset, if changing it.")] string? end,
        [McpToolProperty("location", "New location, if changing it.")] string? location,
        [McpToolProperty("body", "New notes/description, if changing it.")] string? body,
        [McpToolProperty("attendees", "New comma-separated attendee email addresses, if changing it. Replaces the existing attendee list entirely.")]
            string? attendees)
    {
        var attendeeAddresses = string.IsNullOrWhiteSpace(attendees)
            ? null
            : attendees.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

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
