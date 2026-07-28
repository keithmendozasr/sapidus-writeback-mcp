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
            "Edit fields on an existing calendar event. " +
                "If this call fails with an authentication/401-style error, tell the user the outlook-writeback " +
                "connector may need to be reconnected (Settings/Customize > Connectors > outlook-writeback > " +
                "Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("eventId", "The calendar event's ID.", isRequired: true)] string eventId,
        [McpToolProperty("subject", "New subject/title, if changing it.")] string? subject,
        [McpToolProperty("start", "New start time, ISO 8601 with a timezone offset, if changing it.")] string? start,
        [McpToolProperty("end", "New end time, ISO 8601 with a timezone offset, if changing it.")] string? end,
        [McpToolProperty(
            "timeZone",
            "New IANA time zone identifier, if changing how start/end are displayed. Must be provided together " +
                "with start and/or end - passing it alone is rejected as an error. If only one of start/end is " +
                "being changed, only that side's timezone updates; the other keeps its previous value.")]
            string? timeZone,
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
            timeZone,
            location,
            body,
            attendeeAddresses);

        return $"Event updated. ID: {updatedId}.";
    }
}
