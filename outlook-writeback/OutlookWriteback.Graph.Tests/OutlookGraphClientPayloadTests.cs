using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Graph;

namespace OutlookWriteback.Graph.Tests;

[TestFixture]
[Category("Unit")]
public class OutlookGraphClientPayloadTests
{
    [Test]
    public void BuildDraftMessage_maps_recipient_subject_and_plain_text_body()
    {
        var message = OutlookGraphClient.BuildDraftMessage(
            "owner@example.com",
            "Quarterly report",
            "Draft body text.");

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.EqualTo("Quarterly report"));
            Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Text));
            Assert.That(message.Body?.Content, Is.EqualTo("Draft body text."));
            Assert.That(message.ToRecipients?.Single().EmailAddress?.Address, Is.EqualTo("owner@example.com"));
        });
    }

    [Test]
    public void BuildDraftMessage_sets_html_content_type_when_isHtml_is_true()
    {
        var message = OutlookGraphClient.BuildDraftMessage(
            "owner@example.com",
            "Metrics for period ending July 8, 2026",
            "<table><tr><td>2026-07-08</td></tr></table>",
            isHtml: true);

        Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Html));
    }

    [Test]
    public void BuildEvent_converts_to_the_provided_timeZones_local_wall_clock_and_omits_optional_fields_when_absent()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));
        var end = start.AddHours(1);

        var calendarEvent = OutlookGraphClient.BuildEvent(
            "Standup", start, end, timeZone: "America/New_York", location: null, bodyText: null);

        var expectedLocal = TimeZoneInfo.ConvertTime(start, TimeZoneInfo.FindSystemTimeZoneById("America/New_York"));

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.EqualTo("Standup"));
            Assert.That(calendarEvent.Start?.TimeZone, Is.EqualTo("America/New_York"));
            Assert.That(calendarEvent.Start?.DateTime, Is.EqualTo(expectedLocal.DateTime.ToString("o")));
            Assert.That(calendarEvent.Start?.DateTime, Does.Not.EndWith("Z"));
            Assert.That(calendarEvent.Location, Is.Null);
            Assert.That(calendarEvent.Body, Is.Null);
        });
    }

    [Test]
    public void BuildEvent_converts_local_wall_clock_correctly_across_a_DST_boundary()
    {
        var summerStart = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-4)); // EDT
        var winterStart = new DateTimeOffset(2026, 1, 15, 9, 0, 0, TimeSpan.FromHours(-5)); // EST
        var timeZone = TimeZoneInfo.FindSystemTimeZoneById("America/New_York");

        var summerEvent = OutlookGraphClient.BuildEvent(
            "Summer", summerStart, summerStart.AddHours(1), timeZone: "America/New_York", location: null, bodyText: null);
        var winterEvent = OutlookGraphClient.BuildEvent(
            "Winter", winterStart, winterStart.AddHours(1), timeZone: "America/New_York", location: null, bodyText: null);

        Assert.Multiple(() =>
        {
            Assert.That(
                summerEvent.Start?.DateTime,
                Is.EqualTo(TimeZoneInfo.ConvertTime(summerStart, timeZone).DateTime.ToString("o")));
            Assert.That(
                winterEvent.Start?.DateTime,
                Is.EqualTo(TimeZoneInfo.ConvertTime(winterStart, timeZone).DateTime.ToString("o")));
        });
    }

    [Test]
    public void BuildEvent_uses_no_offset_suffix_even_when_timeZone_is_explicitly_UTC()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        var calendarEvent = OutlookGraphClient.BuildEvent(
            "Standup", start, start.AddHours(1), timeZone: "UTC", location: null, bodyText: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Start?.TimeZone, Is.EqualTo("UTC"));
            Assert.That(calendarEvent.Start?.DateTime, Does.Not.EndWith("Z"));
            Assert.That(calendarEvent.Start?.DateTime, Is.Not.EqualTo(start.UtcDateTime.ToString("o")));
        });
    }

    [Test]
    public void BuildEvent_throws_when_timeZone_is_invalid()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        Assert.That(
            () => OutlookGraphClient.BuildEvent("Standup", start, start.AddHours(1), timeZone: "Not/AZone", location: null, bodyText: null),
            Throws.TypeOf<TimeZoneNotFoundException>());
    }

    [Test]
    public void BuildEvent_maps_attendees_when_provided()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        var calendarEvent = OutlookGraphClient.BuildEvent(
            "Standup",
            start,
            start.AddHours(1),
            timeZone: "UTC",
            location: null,
            bodyText: null,
            attendeeAddresses: ["alice@example.com", "bob@example.com"]);

        Assert.That(
            calendarEvent.Attendees?.Select(attendee => attendee.EmailAddress?.Address),
            Is.EqualTo(new[] { "alice@example.com", "bob@example.com" }));
    }

    [Test]
    public void BuildEvent_sets_reminder_fields_when_reminderMinutesBeforeStart_is_provided()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        var calendarEvent = OutlookGraphClient.BuildEvent(
            "Standup", start, start.AddHours(1), timeZone: "UTC", location: null, bodyText: null, reminderMinutesBeforeStart: 15);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.IsReminderOn, Is.True);
            Assert.That(calendarEvent.ReminderMinutesBeforeStart, Is.EqualTo(15));
        });
    }

    [Test]
    public void BuildEvent_leaves_reminder_fields_null_when_reminderMinutesBeforeStart_is_omitted()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        var calendarEvent = OutlookGraphClient.BuildEvent(
            "Standup", start, start.AddHours(1), timeZone: "UTC", location: null, bodyText: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.IsReminderOn, Is.Null);
            Assert.That(calendarEvent.ReminderMinutesBeforeStart, Is.Null);
        });
    }

    [Test]
    public void BuildEvent_accepts_a_zero_minute_reminder()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        var calendarEvent = OutlookGraphClient.BuildEvent(
            "Standup", start, start.AddHours(1), timeZone: "UTC", location: null, bodyText: null, reminderMinutesBeforeStart: 0);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.IsReminderOn, Is.True);
            Assert.That(calendarEvent.ReminderMinutesBeforeStart, Is.EqualTo(0));
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_only_the_subject_when_only_subject_changes()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(toAddress: null, subject: "New subject", bodyText: null);

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.EqualTo("New subject"));
            Assert.That(message.Body, Is.Null);
            Assert.That(message.ToRecipients, Is.Null);
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_only_the_body_when_only_body_changes()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(toAddress: null, subject: null, bodyText: "New body text.");

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.Null);
            Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Text));
            Assert.That(message.Body?.Content, Is.EqualTo("New body text."));
            Assert.That(message.ToRecipients, Is.Null);
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_only_the_recipient_when_only_recipient_changes()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(toAddress: "owner@example.com", subject: null, bodyText: null);

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.Null);
            Assert.That(message.Body, Is.Null);
            Assert.That(message.ToRecipients?.Single().EmailAddress?.Address, Is.EqualTo("owner@example.com"));
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_all_fields_when_all_are_provided()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(
            toAddress: "owner@example.com",
            subject: "New subject",
            bodyText: "New body text.");

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.EqualTo("New subject"));
            Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Text));
            Assert.That(message.Body?.Content, Is.EqualTo("New body text."));
            Assert.That(message.ToRecipients?.Single().EmailAddress?.Address, Is.EqualTo("owner@example.com"));
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_html_content_type_when_isHtml_is_true()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(
            toAddress: null,
            subject: null,
            bodyText: "<table><tr><td>2026-07-08</td></tr></table>",
            isHtml: true);

        Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Html));
    }

    [Test]
    public void BuildUpdateEvent_sets_only_the_subject_when_only_subject_changes()
    {
        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: "New subject",
            start: null,
            end: null,
            timeZone: null,
            location: null,
            bodyText: null,
            attendeeAddresses: null,
            reminderMinutesBeforeStart: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.EqualTo("New subject"));
            Assert.That(calendarEvent.Start, Is.Null);
            Assert.That(calendarEvent.End, Is.Null);
            Assert.That(calendarEvent.Location, Is.Null);
            Assert.That(calendarEvent.Body, Is.Null);
            Assert.That(calendarEvent.Attendees, Is.Null);
        });
    }

    [Test]
    public void BuildUpdateEvent_sets_only_the_body_when_only_body_changes()
    {
        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: null,
            end: null,
            timeZone: null,
            location: null,
            bodyText: "New notes.",
            attendeeAddresses: null,
            reminderMinutesBeforeStart: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.Null);
            Assert.That(calendarEvent.Body?.ContentType, Is.EqualTo(BodyType.Text));
            Assert.That(calendarEvent.Body?.Content, Is.EqualTo("New notes."));
        });
    }

    [Test]
    public void BuildUpdateEvent_sets_only_start_and_end_when_only_the_time_changes_and_preserves_legacy_utc_with_Z_when_timeZone_is_not_provided()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));
        var end = start.AddHours(1);

        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: start,
            end: end,
            timeZone: null,
            location: null,
            bodyText: null,
            attendeeAddresses: null,
            reminderMinutesBeforeStart: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.Null);
            Assert.That(calendarEvent.Start?.TimeZone, Is.EqualTo("UTC"));
            Assert.That(calendarEvent.Start?.DateTime, Is.EqualTo(start.UtcDateTime.ToString("o")));
            Assert.That(calendarEvent.End?.DateTime, Is.EqualTo(end.UtcDateTime.ToString("o")));
            Assert.That(calendarEvent.Location, Is.Null);
        });
    }

    [Test]
    public void BuildUpdateEvent_converts_local_wall_clock_when_timeZone_is_provided_with_start_and_end()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));
        var end = start.AddHours(1);

        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: start,
            end: end,
            timeZone: "America/New_York",
            location: null,
            bodyText: null,
            attendeeAddresses: null,
            reminderMinutesBeforeStart: null);

        var expectedLocal = TimeZoneInfo.ConvertTime(start, TimeZoneInfo.FindSystemTimeZoneById("America/New_York"));

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Start?.TimeZone, Is.EqualTo("America/New_York"));
            Assert.That(calendarEvent.Start?.DateTime, Is.EqualTo(expectedLocal.DateTime.ToString("o")));
            Assert.That(calendarEvent.Start?.DateTime, Does.Not.EndWith("Z"));
        });
    }

    [Test]
    public void BuildUpdateEvent_sets_only_the_location_when_only_location_changes()
    {
        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: null,
            end: null,
            timeZone: null,
            location: "Conference Room B",
            bodyText: null,
            attendeeAddresses: null,
            reminderMinutesBeforeStart: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.Null);
            Assert.That(calendarEvent.Location?.DisplayName, Is.EqualTo("Conference Room B"));
            Assert.That(calendarEvent.Start, Is.Null);
        });
    }

    [Test]
    public void BuildUpdateEvent_sets_only_the_attendees_when_only_attendees_change()
    {
        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: null,
            end: null,
            timeZone: null,
            location: null,
            bodyText: null,
            attendeeAddresses: ["alice@example.com"],
            reminderMinutesBeforeStart: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.Null);
            Assert.That(
                calendarEvent.Attendees?.Select(attendee => attendee.EmailAddress?.Address),
                Is.EqualTo(new[] { "alice@example.com" }));
        });
    }

    [Test]
    public void BuildUpdateEvent_sets_only_reminder_fields_when_only_reminder_changes()
    {
        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: null,
            end: null,
            timeZone: null,
            location: null,
            bodyText: null,
            attendeeAddresses: null,
            reminderMinutesBeforeStart: 30);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.Null);
            Assert.That(calendarEvent.Start, Is.Null);
            Assert.That(calendarEvent.IsReminderOn, Is.True);
            Assert.That(calendarEvent.ReminderMinutesBeforeStart, Is.EqualTo(30));
        });
    }

    [Test]
    public void BuildUpdateEvent_sets_all_fields_when_all_are_provided()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));
        var end = start.AddHours(1);

        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: "New subject",
            start: start,
            end: end,
            timeZone: "America/New_York",
            location: "Conference Room B",
            bodyText: "New notes.",
            attendeeAddresses: ["alice@example.com"],
            reminderMinutesBeforeStart: 15);

        var expectedLocal = TimeZoneInfo.ConvertTime(start, TimeZoneInfo.FindSystemTimeZoneById("America/New_York"));

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.EqualTo("New subject"));
            Assert.That(calendarEvent.Start?.TimeZone, Is.EqualTo("America/New_York"));
            Assert.That(calendarEvent.Start?.DateTime, Is.EqualTo(expectedLocal.DateTime.ToString("o")));
            Assert.That(calendarEvent.Start?.DateTime, Does.Not.EndWith("Z"));
            Assert.That(calendarEvent.Location?.DisplayName, Is.EqualTo("Conference Room B"));
            Assert.That(calendarEvent.Body?.Content, Is.EqualTo("New notes."));
            Assert.That(
                calendarEvent.Attendees?.Select(attendee => attendee.EmailAddress?.Address),
                Is.EqualTo(new[] { "alice@example.com" }));
            Assert.That(calendarEvent.IsReminderOn, Is.True);
            Assert.That(calendarEvent.ReminderMinutesBeforeStart, Is.EqualTo(15));
        });
    }

    [Test]
    public void UpdateDraftAsync_throws_when_no_fields_are_provided()
    {
        var httpClient = new HttpClient { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var client = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));

        Assert.That(
            () => client.UpdateDraftAsync("AAMk-fake-draft-id"),
            Throws.ArgumentException);
    }

    [Test]
    public void UpdateEventAsync_throws_when_no_fields_are_provided()
    {
        var httpClient = new HttpClient { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var client = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));

        Assert.That(
            () => client.UpdateEventAsync("AAkA-fake-event-id"),
            Throws.ArgumentException);
    }

    [Test]
    public void UpdateEventAsync_throws_when_timeZone_is_provided_without_start_or_end()
    {
        var httpClient = new HttpClient { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var client = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));

        Assert.That(
            () => client.UpdateEventAsync("AAkA-fake-event-id", timeZone: "America/New_York"),
            Throws.ArgumentException);
    }

    [Test]
    public void UpdateEventAsync_throws_when_reminderMinutesBeforeStart_is_negative()
    {
        var httpClient = new HttpClient { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var client = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));

        Assert.That(
            () => client.UpdateEventAsync("AAkA-fake-event-id", reminderMinutesBeforeStart: -5),
            Throws.ArgumentException);
    }

    [Test]
    public void CreateEventAsync_throws_when_reminderMinutesBeforeStart_is_negative()
    {
        var httpClient = new HttpClient { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var client = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        Assert.That(
            () => client.CreateEventAsync("Standup", start, start.AddHours(1), "UTC", reminderMinutesBeforeStart: -5),
            Throws.ArgumentException);
    }
}
