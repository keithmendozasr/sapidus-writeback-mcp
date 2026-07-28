using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Functions;

public sealed class CreateEventTool(OutlookGraphClient client)
{
    [Function(nameof(CreateEventTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "create_event",
            "Create a calendar event on the user's calendar. attendees entries must be non-blank - a generic failure " +
                "with no specific reason usually means a blank entry slipped into attendees, so re-check it before " +
                "retrying. If this call fails with an authentication/401-style error, tell the user the " +
                "outlook-writeback connector may need to be reconnected (Settings/Customize > Connectors > " +
                "outlook-writeback > Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("subject", "Event subject/title.", isRequired: true)] string subject,
        [McpToolProperty("start", "Event start time, ISO 8601 with a timezone offset (e.g. 2026-08-01T09:00:00-05:00).", isRequired: true)]
            string start,
        [McpToolProperty("end", "Event end time, ISO 8601 with a timezone offset.", isRequired: true)] string end,
        [McpToolProperty("location", "Event location, if any.")] string? location,
        [McpToolProperty("body", "Event notes/description, if any.")] string? body,
        [McpToolProperty("attendees", "Attendee email addresses, if any.")] string[]? attendees)
    {
        var attendeeAddresses = RecipientList.Normalize(attendees);

        var eventId = await client.CreateEventAsync(
            subject,
            DateTimeOffset.Parse(start),
            DateTimeOffset.Parse(end),
            location,
            body,
            attendeeAddresses);

        return $"Event created. ID: {eventId}.";
    }
}
