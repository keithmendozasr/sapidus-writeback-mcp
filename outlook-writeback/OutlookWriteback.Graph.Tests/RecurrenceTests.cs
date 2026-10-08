using Microsoft.Graph.Models;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Graph.Tests;

[TestFixture]
[Category("Unit")]
public class RecurrenceTests
{
    private const string NewYork = "America/New_York";

    // 2026-08-04 is the first Tuesday of August 2026 (and a Tuesday); 2026-08-25 is the last Tuesday;
    // 2026-09-04 is the first Friday of September 2026.
    private static Event Build(RecurrenceSpec spec, DateTimeOffset start, string timeZone = NewYork) =>
        OutlookGraphClient.BuildEvent("PT", start, start.AddHours(1), timeZone, null, null, recurrence: spec);

    private static DateTimeOffset Local(int y, int m, int d, int hour = 9) =>
        new(y, m, d, hour, 0, 0, TimeZoneInfo.FindSystemTimeZoneById(NewYork).GetUtcOffset(new DateTime(y, m, d, hour, 0, 0)));

    [Test]
    public void BuildEvent_leaves_Recurrence_null_when_no_recurrence_is_given()
    {
        var start = Local(2026, 8, 4);
        var calendarEvent = OutlookGraphClient.BuildEvent("One-off", start, start.AddHours(1), NewYork, null, null);

        Assert.That(calendarEvent.Recurrence, Is.Null);
    }

    [Test]
    public void BuildEvent_maps_a_daily_series_with_an_end_date()
    {
        var spec = RecurrenceSpec.Create("daily", 3, null, null, null, null, "2026-08-31");

        var recurrence = Build(spec, Local(2026, 8, 4)).Recurrence!;

        Assert.Multiple(() =>
        {
            Assert.That(recurrence.Pattern!.Type, Is.EqualTo(RecurrencePatternType.Daily));
            Assert.That(recurrence.Pattern.Interval, Is.EqualTo(3));
            Assert.That(recurrence.Pattern.DaysOfWeek, Is.Null);
            Assert.That(recurrence.Pattern.DayOfMonth, Is.Null);
            Assert.That(recurrence.Range!.Type, Is.EqualTo(RecurrenceRangeType.EndDate));
            Assert.That(recurrence.Range.StartDate?.ToString(), Is.EqualTo("2026-08-04"));
            Assert.That(recurrence.Range.EndDate?.ToString(), Is.EqualTo("2026-08-31"));
            Assert.That(recurrence.Range.RecurrenceTimeZone, Is.EqualTo(NewYork));
        });
    }

    [Test]
    public void BuildEvent_maps_a_weekly_series_with_multiple_days_and_a_count()
    {
        var spec = RecurrenceSpec.Create("Weekly", null, ["Tuesday", "thursday", "TUESDAY"], null, null, 8, null);

        var recurrence = Build(spec, Local(2026, 8, 4)).Recurrence!;

        Assert.Multiple(() =>
        {
            Assert.That(recurrence.Pattern!.Type, Is.EqualTo(RecurrencePatternType.Weekly));
            Assert.That(recurrence.Pattern.Interval, Is.EqualTo(1));
            Assert.That(recurrence.Pattern.DaysOfWeek, Is.EqualTo(new List<DayOfWeekObject?> { DayOfWeekObject.Tuesday, DayOfWeekObject.Thursday }));
            Assert.That(recurrence.Pattern.FirstDayOfWeek, Is.EqualTo(DayOfWeekObject.Sunday));
            Assert.That(recurrence.Range!.Type, Is.EqualTo(RecurrenceRangeType.Numbered));
            Assert.That(recurrence.Range.NumberOfOccurrences, Is.EqualTo(8));
            Assert.That(recurrence.Range.EndDate, Is.Null);
        });
    }

    [Test]
    public void BuildEvent_maps_an_absolute_monthly_series_with_no_end_using_the_local_start_day()
    {
        var spec = RecurrenceSpec.Create("monthly", 2, null, null, null, null, null);

        var recurrence = Build(spec, Local(2026, 8, 15)).Recurrence!;

        Assert.Multiple(() =>
        {
            Assert.That(recurrence.Pattern!.Type, Is.EqualTo(RecurrencePatternType.AbsoluteMonthly));
            Assert.That(recurrence.Pattern.DayOfMonth, Is.EqualTo(15));
            Assert.That(recurrence.Pattern.Interval, Is.EqualTo(2));
            Assert.That(recurrence.Pattern.Index, Is.Null);
            Assert.That(recurrence.Range!.Type, Is.EqualTo(RecurrenceRangeType.NoEnd));
        });
    }

    [Test]
    public void BuildEvent_maps_a_relative_monthly_series_second_tuesday()
    {
        var spec = RecurrenceSpec.Create("monthly", null, ["tuesday"], "second", null, null, null);

        var recurrence = Build(spec, Local(2026, 8, 11)).Recurrence!; // 2nd Tuesday of Aug 2026

        Assert.Multiple(() =>
        {
            Assert.That(recurrence.Pattern!.Type, Is.EqualTo(RecurrencePatternType.RelativeMonthly));
            Assert.That(recurrence.Pattern.Index, Is.EqualTo(WeekIndex.Second));
            Assert.That(recurrence.Pattern.DaysOfWeek, Is.EqualTo(new List<DayOfWeekObject?> { DayOfWeekObject.Tuesday }));
            Assert.That(recurrence.Pattern.DayOfMonth, Is.Null);
        });
    }

    [Test]
    public void BuildEvent_maps_a_relative_monthly_series_last_tuesday()
    {
        var spec = RecurrenceSpec.Create("monthly", null, ["tuesday"], "last", null, null, null);

        var recurrence = Build(spec, Local(2026, 8, 25)).Recurrence!; // last Tuesday of Aug 2026

        Assert.That(recurrence.Pattern!.Index, Is.EqualTo(WeekIndex.Last));
    }

    [Test]
    public void BuildEvent_maps_an_absolute_yearly_series_from_the_start_date()
    {
        var spec = RecurrenceSpec.Create("yearly", null, null, null, null, 5, null);

        var recurrence = Build(spec, Local(2026, 3, 15)).Recurrence!;

        Assert.Multiple(() =>
        {
            Assert.That(recurrence.Pattern!.Type, Is.EqualTo(RecurrencePatternType.AbsoluteYearly));
            Assert.That(recurrence.Pattern.DayOfMonth, Is.EqualTo(15));
            Assert.That(recurrence.Pattern.Month, Is.EqualTo(3));
        });
    }

    [Test]
    public void BuildEvent_maps_a_relative_yearly_series_first_friday_in_september()
    {
        var spec = RecurrenceSpec.Create("yearly", null, ["friday"], "first", 9, null, null);

        var recurrence = Build(spec, Local(2026, 9, 4)).Recurrence!;

        Assert.Multiple(() =>
        {
            Assert.That(recurrence.Pattern!.Type, Is.EqualTo(RecurrencePatternType.RelativeYearly));
            Assert.That(recurrence.Pattern.Index, Is.EqualTo(WeekIndex.First));
            Assert.That(recurrence.Pattern.Month, Is.EqualTo(9));
            Assert.That(recurrence.Pattern.DayOfMonth, Is.Null);
        });
    }

    [Test]
    public void BuildEvent_defaults_a_relative_yearly_month_to_the_start_month()
    {
        var spec = RecurrenceSpec.Create("yearly", null, ["friday"], "first", null, null, null);

        var recurrence = Build(spec, Local(2026, 9, 4)).Recurrence!;

        Assert.That(recurrence.Pattern!.Month, Is.EqualTo(9));
    }

    [Test]
    public void BuildEvent_takes_the_range_start_date_from_the_local_date_not_the_UTC_date()
    {
        // 9pm on Aug 3 in New York is Aug 4 01:00 UTC; the series must start on the local Aug 3 (a Monday).
        var start = new DateTimeOffset(2026, 8, 4, 1, 0, 0, TimeSpan.Zero);
        var spec = RecurrenceSpec.Create("weekly", null, ["monday"], null, null, 4, null);

        var recurrence = Build(spec, start).Recurrence!;

        Assert.That(recurrence.Range!.StartDate?.ToString(), Is.EqualTo("2026-08-03"));
    }

    [TestCase("hourly", null, null, null, null, null, null, "recurrenceType")]
    [TestCase("daily", 0, null, null, null, null, null, "recurrenceInterval")]
    [TestCase("daily", null, null, null, null, 0, null, "recurrenceCount")]
    [TestCase("daily", null, null, null, null, 3, "2026-09-01", "not both")]
    [TestCase("daily", null, null, null, null, null, "09/01/2026", "yyyy-MM-dd")]
    [TestCase("weekly", null, null, null, null, null, null, "recurrenceDaysOfWeek is required")]
    [TestCase("weekly", null, new[] { "funday" }, null, null, null, null, "day names")]
    [TestCase("daily", null, new[] { "monday" }, null, null, null, null, "recurrenceDaysOfWeek")]
    [TestCase("weekly", null, new[] { "monday" }, "first", null, null, null, "recurrenceWeekIndex")]
    [TestCase("monthly", null, new[] { "monday" }, null, null, null, null, "requires recurrenceWeekIndex")]
    [TestCase("monthly", null, new[] { "monday", "tuesday" }, "first", null, null, null, "exactly one")]
    [TestCase("monthly", null, new[] { "monday" }, "fifth", null, null, null, "recurrenceWeekIndex")]
    [TestCase("monthly", null, new[] { "monday" }, "first", 9, null, null, "yearly")]
    [TestCase("yearly", null, new[] { "monday" }, "first", 13, null, null, "between 1 and 12")]
    [TestCase("yearly", null, null, null, 9, null, null, "requires recurrenceWeekIndex")]
    public void Create_rejects_invalid_input(
        string type, int? interval, string[]? days, string? index, int? month, int? count, string? until, string messagePart)
    {
        Assert.That(
            () => RecurrenceSpec.Create(type, interval, days, index, month, count, until),
            Throws.ArgumentException.With.Message.Contain(messagePart));
    }

    [Test]
    public void BuildEvent_rejects_an_until_date_before_the_start()
    {
        var spec = RecurrenceSpec.Create("daily", null, null, null, null, null, "2026-08-01");

        Assert.That(() => Build(spec, Local(2026, 8, 4)), Throws.ArgumentException.With.Message.Contain("before the event's start"));
    }

    [Test]
    public void BuildEvent_rejects_a_weekly_series_whose_start_day_is_not_in_the_list()
    {
        var spec = RecurrenceSpec.Create("weekly", null, ["wednesday"], null, null, null, null);

        Assert.That(() => Build(spec, Local(2026, 8, 4)), Throws.ArgumentException.With.Message.Contain("isn't in recurrenceDaysOfWeek"));
    }

    [Test]
    public void BuildEvent_rejects_a_relative_monthly_series_whose_start_is_not_that_weekday_of_the_month()
    {
        var spec = RecurrenceSpec.Create("monthly", null, ["tuesday"], "second", null, null, null);

        // Aug 4 is the FIRST Tuesday.
        Assert.That(() => Build(spec, Local(2026, 8, 4)), Throws.ArgumentException.With.Message.Contain("second Tuesday"));
    }

    [Test]
    public void BuildEvent_rejects_a_relative_yearly_series_whose_start_is_in_the_wrong_month()
    {
        var spec = RecurrenceSpec.Create("yearly", null, ["friday"], "first", 9, null, null);

        // Aug 7 2026 is the first Friday of August, not September.
        Assert.That(() => Build(spec, Local(2026, 8, 7)), Throws.ArgumentException.With.Message.Contain("recurrenceMonth"));
    }

    // ---- update_event ----

    private static Event BuildUpdate(RecurrenceSpec? spec, DateTimeOffset? start, string? timeZone = NewYork) =>
        OutlookGraphClient.BuildUpdateEvent(null, start, start?.AddHours(1), timeZone, null, null, null, null, spec);

    private static OutlookGraphClient ClientThatMustNotCallGraph() =>
        new(new Microsoft.Graph.GraphServiceClient(
            new HttpClient { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") },
            new Microsoft.Kiota.Abstractions.Authentication.AnonymousAuthenticationProvider()));

    [Test]
    public void BuildUpdateEvent_leaves_Recurrence_null_when_no_recurrence_is_given()
    {
        Assert.That(BuildUpdate(null, Local(2026, 8, 4)).Recurrence, Is.Null);
    }

    [Test]
    public void BuildUpdateEvent_maps_a_weekly_recurrence_alongside_start_and_end()
    {
        var spec = RecurrenceSpec.Create("weekly", null, ["tuesday"], null, null, 12, null);

        var calendarEvent = BuildUpdate(spec, Local(2026, 8, 4));

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Start?.TimeZone, Is.EqualTo(NewYork));
            Assert.That(calendarEvent.Recurrence!.Pattern!.Type, Is.EqualTo(RecurrencePatternType.Weekly));
            Assert.That(calendarEvent.Recurrence.Range!.StartDate?.ToString(), Is.EqualTo("2026-08-04"));
            Assert.That(calendarEvent.Recurrence.Range.NumberOfOccurrences, Is.EqualTo(12));
            Assert.That(calendarEvent.Subject, Is.Null);
        });
    }

    [Test]
    public void BuildUpdateEvent_maps_a_relative_monthly_recurrence()
    {
        var spec = RecurrenceSpec.Create("monthly", null, ["tuesday"], "second", null, null, null);

        var recurrence = BuildUpdate(spec, Local(2026, 8, 11)).Recurrence!;

        Assert.Multiple(() =>
        {
            Assert.That(recurrence.Pattern!.Type, Is.EqualTo(RecurrencePatternType.RelativeMonthly));
            Assert.That(recurrence.Pattern.Index, Is.EqualTo(WeekIndex.Second));
        });
    }

    [Test]
    public void BuildUpdateEvent_rejects_a_start_that_is_not_part_of_the_series()
    {
        var spec = RecurrenceSpec.Create("weekly", null, ["wednesday"], null, null, null, null);

        Assert.That(() => BuildUpdate(spec, Local(2026, 8, 4)), Throws.ArgumentException.With.Message.Contain("isn't in recurrenceDaysOfWeek"));
    }

    [Test]
    public void BuildUpdateEvent_rejects_a_recurrence_without_start_or_timeZone()
    {
        var spec = RecurrenceSpec.Create("daily", null, null, null, null, null, null);

        Assert.Multiple(() =>
        {
            Assert.That(() => BuildUpdate(spec, null), Throws.ArgumentException);
            Assert.That(() => BuildUpdate(spec, Local(2026, 8, 4), timeZone: null), Throws.ArgumentException);
        });
    }

    [Test]
    public void UpdateEventAsync_requires_start_when_recurrence_is_provided()
    {
        var spec = RecurrenceSpec.Create("daily", null, null, null, null, null, null);

        Assert.That(
            () => ClientThatMustNotCallGraph().UpdateEventAsync("AAkA-fake-event-id", timeZone: NewYork, recurrence: spec),
            Throws.ArgumentException);
    }

    [Test]
    public void UpdateEventAsync_requires_timeZone_when_recurrence_is_provided()
    {
        var spec = RecurrenceSpec.Create("daily", null, null, null, null, null, null);

        Assert.That(
            () => ClientThatMustNotCallGraph().UpdateEventAsync("AAkA-fake-event-id", start: Local(2026, 8, 4), recurrence: spec),
            Throws.ArgumentException.With.Message.Contain("start and timeZone"));
    }

    [Test]
    public void CreateOptional_returns_null_when_no_recurrence_input_is_given()
    {
        Assert.That(RecurrenceSpec.CreateOptional(null, null, null, null, null, null, null), Is.Null);
    }

    [Test]
    public void CreateOptional_throws_when_other_recurrence_input_arrives_without_a_type()
    {
        Assert.That(
            () => RecurrenceSpec.CreateOptional(null, 2, null, null, null, null, null),
            Throws.ArgumentException.With.Message.Contain("recurrenceType is required"));
    }

    [Test]
    public void CreateOptional_builds_a_spec_when_a_type_is_given()
    {
        var spec = RecurrenceSpec.CreateOptional("daily", 2, null, null, null, null, null);

        Assert.That(spec, Is.Not.Null.And.Property("Interval").EqualTo(2));
    }
}
