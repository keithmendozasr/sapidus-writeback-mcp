using System.Globalization;

namespace OutlookWriteback.Graph;

public enum RecurrenceFrequency
{
    Daily,
    Weekly,
    Monthly,
    Yearly,
}

public enum RecurrenceWeekIndex
{
    First,
    Second,
    Third,
    Fourth,
    Last,
}

/// <summary>
/// A validated, Graph-agnostic description of how a new calendar event repeats. Built from the raw
/// create_event tool strings by <see cref="Create"/>, which checks everything that doesn't depend on
/// the event's start date; <see cref="ValidateAgainstStart"/> then checks the pieces that do.
/// Monthly/yearly are "relative" (e.g. second Tuesday) when <see cref="WeekIndex"/> is set, otherwise
/// "absolute" (a fixed day of the month, taken from the event's start date).
/// </summary>
public sealed record RecurrenceSpec(
    RecurrenceFrequency Frequency,
    int Interval,
    IReadOnlyList<DayOfWeek> DaysOfWeek,
    RecurrenceWeekIndex? WeekIndex,
    int? Month,
    int? Count,
    DateOnly? Until)
{
    public bool IsRelative => WeekIndex is not null;

    /// <summary>
    /// For tools where recurrence is optional: null when no recurrence* input was given, a validated spec when
    /// <paramref name="type"/> is, and an error when other recurrence* inputs arrive without a type.
    /// </summary>
    public static RecurrenceSpec? CreateOptional(
        string? type,
        int? interval,
        IEnumerable<string>? daysOfWeek,
        string? weekIndex,
        int? month,
        int? count,
        string? until)
    {
        if (type is not null)
            return Create(type, interval, daysOfWeek, weekIndex, month, count, until);

        if (interval is not null || daysOfWeek is not null || weekIndex is not null || month is not null || count is not null || until is not null)
            throw new ArgumentException("recurrenceType is required when any other recurrence* parameter is provided.");

        return null;
    }

    public static RecurrenceSpec Create(
        string type,
        int? interval,
        IEnumerable<string>? daysOfWeek,
        string? weekIndex,
        int? month,
        int? count,
        string? until)
    {
        if (!Enum.TryParse<RecurrenceFrequency>(type?.Trim(), ignoreCase: true, out var frequency)
            || !Enum.IsDefined(frequency))
        {
            throw new ArgumentException($"recurrenceType must be one of daily, weekly, monthly, yearly (got '{type}').");
        }

        if (interval is < 1)
            throw new ArgumentException("recurrenceInterval must be at least 1.");

        if (count is < 1)
            throw new ArgumentException("recurrenceCount must be at least 1.");

        if (count is not null && until is not null)
            throw new ArgumentException("Provide recurrenceCount or recurrenceUntil, not both.");

        DateOnly? untilDate = null;
        if (until is not null)
        {
            if (!DateOnly.TryParseExact(until.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            {
                throw new ArgumentException(
                    $"recurrenceUntil must be a date formatted exactly yyyy-MM-dd, e.g. 2027-04-09 - no time, no offset, " +
                    $"and not month/day/year order (got '{until}').");
            }

            untilDate = parsed;
        }

        var days = ParseDays(daysOfWeek);

        RecurrenceWeekIndex? index = null;
        if (weekIndex is not null)
        {
            if (!Enum.TryParse<RecurrenceWeekIndex>(weekIndex.Trim(), ignoreCase: true, out var parsedIndex)
                || !Enum.IsDefined(parsedIndex))
            {
                throw new ArgumentException($"recurrenceWeekIndex must be one of first, second, third, fourth, last (got '{weekIndex}').");
            }

            index = parsedIndex;
        }

        if (month is not null && (month < 1 || month > 12))
            throw new ArgumentException("recurrenceMonth must be between 1 and 12.");

        switch (frequency)
        {
            case RecurrenceFrequency.Daily:
                RejectIf(days.Count > 0, "recurrenceDaysOfWeek only applies to weekly recurrence, or monthly/yearly recurrence with recurrenceWeekIndex.");
                RejectIf(index is not null, "recurrenceWeekIndex only applies to monthly or yearly recurrence.");
                RejectIf(month is not null, "recurrenceMonth only applies to yearly recurrence with recurrenceWeekIndex.");
                break;

            case RecurrenceFrequency.Weekly:
                if (days.Count == 0)
                    throw new ArgumentException("recurrenceDaysOfWeek is required for weekly recurrence.");
                RejectIf(index is not null, "recurrenceWeekIndex only applies to monthly or yearly recurrence.");
                RejectIf(month is not null, "recurrenceMonth only applies to yearly recurrence with recurrenceWeekIndex.");
                break;

            case RecurrenceFrequency.Monthly:
                ValidateMonthlyOrYearly(days, index, month, yearly: false);
                break;

            case RecurrenceFrequency.Yearly:
                ValidateMonthlyOrYearly(days, index, month, yearly: true);
                break;
        }

        return new RecurrenceSpec(frequency, interval ?? 1, days, index, month, count, untilDate);
    }

    /// <summary>
    /// Checks the series against the event's first occurrence (<paramref name="localStart"/>, the start's
    /// date on the calendar in the event's time zone), so the first occurrence is actually part of the
    /// series instead of Graph silently shifting it.
    /// </summary>
    public void ValidateAgainstStart(DateOnly localStart)
    {
        if (Until is not null && Until < localStart)
            throw new ArgumentException($"recurrenceUntil ({Until:yyyy-MM-dd}) must not be before the event's start date ({localStart:yyyy-MM-dd}).");

        if (Frequency == RecurrenceFrequency.Weekly && !DaysOfWeek.Contains(localStart.DayOfWeek))
        {
            throw new ArgumentException(
                $"The event starts on a {localStart.DayOfWeek}, which isn't in recurrenceDaysOfWeek " +
                $"({string.Join(", ", DaysOfWeek)}). Start the event on one of those days.");
        }

        if (!IsRelative)
            return;

        if (Frequency == RecurrenceFrequency.Yearly && localStart.Month != EffectiveMonth(localStart))
        {
            throw new ArgumentException(
                $"The event starts in month {localStart.Month}, but recurrenceMonth is {Month}. " +
                "Start the event in the recurrence month.");
        }

        if (localStart.DayOfWeek != DaysOfWeek[0] || !MatchesWeekIndex(localStart, WeekIndex!.Value))
        {
            throw new ArgumentException(
                $"The event starts on {localStart:yyyy-MM-dd}, which is not the {WeekIndex!.Value.ToString().ToLowerInvariant()} " +
                $"{DaysOfWeek[0]} of its month. Start the event on that day so the first occurrence matches the series.");
        }
    }

    /// <summary>The month a yearly series repeats in: the explicit recurrenceMonth, else the start's month.</summary>
    public int EffectiveMonth(DateOnly localStart) => Month ?? localStart.Month;

    private static void ValidateMonthlyOrYearly(IReadOnlyList<DayOfWeek> days, RecurrenceWeekIndex? index, int? month, bool yearly)
    {
        if (index is null)
        {
            RejectIf(days.Count > 0, "recurrenceDaysOfWeek on a monthly/yearly series requires recurrenceWeekIndex (e.g. second Tuesday); " +
                "without it the series repeats on the start date's day of the month.");
            RejectIf(month is not null, "recurrenceMonth requires recurrenceWeekIndex; without it a yearly series repeats on the start date's month and day.");
            return;
        }

        if (days.Count != 1)
            throw new ArgumentException("A relative monthly/yearly series (with recurrenceWeekIndex) needs exactly one recurrenceDaysOfWeek entry.");

        RejectIf(!yearly && month is not null, "recurrenceMonth only applies to yearly recurrence.");
    }

    private static List<DayOfWeek> ParseDays(IEnumerable<string>? daysOfWeek)
    {
        var days = new List<DayOfWeek>();
        foreach (var raw in daysOfWeek ?? [])
        {
            if (!Enum.TryParse<DayOfWeek>(raw?.Trim(), ignoreCase: true, out var day) || !Enum.IsDefined(day))
                throw new ArgumentException($"recurrenceDaysOfWeek entries must be day names like 'monday' (got '{raw}').");

            if (!days.Contains(day))
                days.Add(day);
        }

        return days;
    }

    private static bool MatchesWeekIndex(DateOnly date, RecurrenceWeekIndex index) => index switch
    {
        RecurrenceWeekIndex.Last => date.Day + 7 > DateTime.DaysInMonth(date.Year, date.Month),
        _ => (date.Day - 1) / 7 == (int)index,
    };

    private static void RejectIf(bool condition, string message)
    {
        if (condition)
            throw new ArgumentException(message);
    }
}
