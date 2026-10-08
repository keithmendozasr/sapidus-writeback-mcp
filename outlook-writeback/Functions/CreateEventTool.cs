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
            "Create a calendar event on the user's calendar, optionally as a recurring series (set recurrenceType plus " +
                "the other recurrence* parameters; omit them all for a single event, and start/end then describe the " +
                "first occurrence). attendees entries must be non-blank - a generic failure " +
                "with no specific reason usually means a blank entry slipped into attendees, so re-check it before " +
                "retrying. If this call fails with an authentication/401-style error, tell the user the " +
                "outlook-writeback connector may need to be reconnected (Settings/Customize > Connectors > " +
                "outlook-writeback > Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("subject", "Event subject/title.", isRequired: true)] string subject,
        [McpToolProperty("start", "Event start time, ISO 8601 with a timezone offset (e.g. 2026-08-01T09:00:00-05:00).", isRequired: true)]
            string start,
        [McpToolProperty("end", "Event end time, ISO 8601 with a timezone offset.", isRequired: true)] string end,
        [McpToolProperty(
            "timeZone",
            "IANA time zone identifier for start/end (e.g. \"America/New_York\"). Windows time zone names are also " +
                "accepted. Controls how the event's time is displayed on the calendar.",
            isRequired: true)]
            string timeZone,
        [McpToolProperty("location", "Event location, if any.")] string? location,
        [McpToolProperty("body", "Event notes/description, if any.")] string? body,
        [McpToolProperty("attendees", "Attendee email addresses, if any.")] string[]? attendees,
        [McpToolProperty(
            "reminderMinutes",
            "Minutes before the event start to show a reminder, if setting one (e.g. 15; 0 means at start time). " +
                "Omit to leave reminders at the mailbox/Graph default.")]
            int? reminderMinutes,
        [McpToolProperty("recurrenceType", RecurrenceParameterDocs.Type)] string? recurrenceType,
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

        var eventId = await client.CreateEventAsync(
            subject,
            DateTimeOffset.Parse(start),
            DateTimeOffset.Parse(end),
            timeZone,
            location,
            body,
            attendeeAddresses,
            reminderMinutes,
            recurrence);

        return recurrence is null ? $"Event created. ID: {eventId}." : $"Recurring event series created. ID: {eventId}.";
    }
}
