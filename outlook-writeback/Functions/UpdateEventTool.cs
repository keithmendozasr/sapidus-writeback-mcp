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
            "Edit fields on an existing calendar event, including changing or adding recurrence (see the recurrence* parameters; removing recurrence is not supported). attendees entries must be non-blank - a generic failure with no " +
                "specific reason usually means a blank entry slipped into attendees, so re-check it before retrying. " +
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
        [McpToolProperty(
            "attendees",
            "New attendee email addresses, if changing them. Omit to leave the attendee list unchanged; pass an empty array to clear it entirely.")]
            string[]? attendees,
        [McpToolProperty(
            "reminderMinutes",
            "Minutes before the event start to show a reminder, if setting/changing one (e.g. 15; 0 means at start " +
                "time). Omit to leave the event's existing reminder state unchanged. There is currently no way to " +
                "explicitly turn off an existing reminder through this tool.")]
            int? reminderMinutes,
        [McpToolProperty("recurrenceType", RecurrenceParameterDocs.Type + " On update this replaces the event's recurrence and requires start and timeZone: set start to the series' first occurrence (its current start, from the M365 connector, if the schedule isn't moving).")]
            string? recurrenceType,
        [McpToolProperty("recurrenceInterval", RecurrenceParameterDocs.Interval)] int? recurrenceInterval,
        [McpToolProperty("recurrenceDaysOfWeek", RecurrenceParameterDocs.DaysOfWeek)] string[]? recurrenceDaysOfWeek,
        [McpToolProperty("recurrenceWeekIndex", RecurrenceParameterDocs.WeekIndex)] string? recurrenceWeekIndex,
        [McpToolProperty("recurrenceMonth", RecurrenceParameterDocs.Month)] int? recurrenceMonth,
        [McpToolProperty("recurrenceCount", RecurrenceParameterDocs.Count)] int? recurrenceCount,
        [McpToolProperty("recurrenceUntil", RecurrenceParameterDocs.Until)] string? recurrenceUntil)
    {
        var attendeeAddresses = RecipientList.Normalize(attendees);
        var recurrence = RecurrenceSpec.CreateOptional(
            recurrenceType, recurrenceInterval, recurrenceDaysOfWeek, recurrenceWeekIndex, recurrenceMonth, recurrenceCount, recurrenceUntil);

        var updatedId = await client.UpdateEventAsync(
            eventId,
            subject,
            start is null ? null : DateTimeOffset.Parse(start),
            end is null ? null : DateTimeOffset.Parse(end),
            timeZone,
            location,
            body,
            attendeeAddresses,
            reminderMinutes,
            recurrence);

        return $"Event updated. ID: {updatedId}.";
    }
}
