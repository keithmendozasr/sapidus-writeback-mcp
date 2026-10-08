namespace OutlookWriteback.Functions;

/// <summary>
/// Tool-parameter descriptions shared by create_event and update_event, so the two tools' MCP schemas
/// can't drift apart. Attribute arguments must be compile-time constants, hence consts.
/// </summary>
internal static class RecurrenceParameterDocs
{
    public const string Type =
        "Makes the event a recurring series: daily, weekly, monthly, or yearly. start is the series' first occurrence. " +
        "Monthly/yearly repeat on the start date's day of the month (yearly: month and day) unless recurrenceWeekIndex " +
        "is set. Omit all recurrence* parameters to not use recurrence.";

    public const string Interval =
        "Repeat every N days/weeks/months/years. Defaults to 1 (e.g. 2 with weekly = every other week).";

    public const string DaysOfWeek =
        "Day names (monday..sunday). Required for weekly (the start date must fall on one of them). For monthly/yearly " +
        "with recurrenceWeekIndex, exactly one day (e.g. [\"tuesday\"]). Not allowed otherwise.";

    public const string WeekIndex =
        "first, second, third, fourth, or last. Only for monthly/yearly: repeat on that weekday of the month, e.g. " +
        "monthly + second + [\"tuesday\"] = second Tuesday of every month. The start date must itself be that weekday.";

    public const string Month =
        "1-12. Only for yearly with recurrenceWeekIndex, e.g. 9 with first + [\"friday\"] = first Friday in September " +
        "every year. Defaults to the start date's month; the start date must be in that month.";

    public const string Count =
        "End the series after this many occurrences. Don't combine with recurrenceUntil; omit both for no end date.";

    public const string Until =
        "End the series on or before this date, formatted exactly yyyy-MM-dd (e.g. 2027-04-09) in the event's timeZone - not M/D/Y. Don't combine with recurrenceCount; omit both for no end date.";
}
