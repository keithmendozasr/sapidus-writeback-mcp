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
        [McpToolProperty(
            "recurrenceType",
            "Makes the event a recurring series: daily, weekly, monthly, or yearly. start/end are the first occurrence. " +
                "Monthly/yearly repeat on the start date's day of the month (yearly: month and day) unless recurrenceWeekIndex " +
                "is set. Omit all recurrence* parameters for a one-off event.")]
            string? recurrenceType,
        [McpToolProperty("recurrenceInterval", "Repeat every N days/weeks/months/years. Defaults to 1 (e.g. 2 with weekly = every other week).")]
            int? recurrenceInterval,
        [McpToolProperty(
            "recurrenceDaysOfWeek",
            "Day names (monday..sunday). Required for weekly (the start date must fall on one of them). For monthly/yearly " +
                "with recurrenceWeekIndex, exactly one day (e.g. [\"tuesday\"]). Not allowed otherwise.")]
            string[]? recurrenceDaysOfWeek,
        [McpToolProperty(
            "recurrenceWeekIndex",
            "first, second, third, fourth, or last. Only for monthly/yearly: repeat on that weekday of the month, e.g. " +
                "monthly + second + [\"tuesday\"] = second Tuesday of every month. The start date must itself be that weekday.")]
            string? recurrenceWeekIndex,
        [McpToolProperty(
            "recurrenceMonth",
            "1-12. Only for yearly with recurrenceWeekIndex, e.g. 9 with first + [\"friday\"] = first Friday in September " +
                "every year. Defaults to the start date's month; the start date must be in that month.")]
            int? recurrenceMonth,
        [McpToolProperty("recurrenceCount", "End the series after this many occurrences. Don't combine with recurrenceUntil; omit both for no end date.")]
            int? recurrenceCount,
        [McpToolProperty(
            "recurrenceUntil",
            "End the series on or before this date, yyyy-MM-dd in the event's timeZone. Don't combine with recurrenceCount; omit both for no end date.")]
            string? recurrenceUntil)
    {
        var attendeeAddresses = RecipientList.Normalize(attendees);
        var recurrence = BuildRecurrence(
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

    private static RecurrenceSpec? BuildRecurrence(
        string? type, int? interval, string[]? daysOfWeek, string? weekIndex, int? month, int? count, string? until)
    {
        if (type is not null)
            return RecurrenceSpec.Create(type, interval, daysOfWeek, weekIndex, month, count, until);

        if (interval is not null || daysOfWeek is not null || weekIndex is not null || month is not null || count is not null || until is not null)
            throw new ArgumentException("recurrenceType is required when any other recurrence* parameter is provided.");

        return null;
    }
}
