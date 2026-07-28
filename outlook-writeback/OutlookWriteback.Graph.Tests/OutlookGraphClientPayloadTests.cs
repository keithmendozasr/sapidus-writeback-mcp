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
            ["owner@example.com"],
            "Quarterly report",
            "Draft body text.");

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.EqualTo("Quarterly report"));
            Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Text));
            Assert.That(message.Body?.Content, Is.EqualTo("Draft body text."));
            Assert.That(
                message.ToRecipients?.Select(recipient => recipient.EmailAddress?.Address),
                Is.EqualTo(new[] { "owner@example.com" }));
        });
    }

    [Test]
    public void BuildDraftMessage_maps_multiple_to_recipients()
    {
        var message = OutlookGraphClient.BuildDraftMessage(
            ["alice@example.com", "bob@example.com"],
            "Quarterly report",
            "Draft body text.");

        Assert.That(
            message.ToRecipients?.Select(recipient => recipient.EmailAddress?.Address),
            Is.EqualTo(new[] { "alice@example.com", "bob@example.com" }));
    }

    [Test]
    public void BuildDraftMessage_maps_cc_and_bcc_when_provided()
    {
        var message = OutlookGraphClient.BuildDraftMessage(
            ["alice@example.com"],
            "Quarterly report",
            "Draft body text.",
            ccAddresses: ["bob@example.com"],
            bccAddresses: ["carol@example.com"]);

        Assert.Multiple(() =>
        {
            Assert.That(
                message.CcRecipients?.Select(recipient => recipient.EmailAddress?.Address),
                Is.EqualTo(new[] { "bob@example.com" }));
            Assert.That(
                message.BccRecipients?.Select(recipient => recipient.EmailAddress?.Address),
                Is.EqualTo(new[] { "carol@example.com" }));
        });
    }

    [Test]
    public void BuildDraftMessage_leaves_all_recipient_lists_empty_when_none_are_provided()
    {
        var message = OutlookGraphClient.BuildDraftMessage(
            null,
            "Placeholder",
            "Draft body text.");

        Assert.Multiple(() =>
        {
            Assert.That(message.ToRecipients, Is.Empty);
            Assert.That(message.CcRecipients, Is.Empty);
            Assert.That(message.BccRecipients, Is.Empty);
        });
    }

    [Test]
    public void BuildDraftMessage_sets_html_content_type_when_isHtml_is_true()
    {
        var message = OutlookGraphClient.BuildDraftMessage(
            ["owner@example.com"],
            "Metrics for period ending July 8, 2026",
            "<table><tr><td>2026-07-08</td></tr></table>",
            isHtml: true);

        Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Html));
    }

    [Test]
    public void BuildEvent_maps_times_to_utc_and_omits_optional_fields_when_absent()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));
        var end = start.AddHours(1);

        var calendarEvent = OutlookGraphClient.BuildEvent("Standup", start, end, location: null, bodyText: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.EqualTo("Standup"));
            Assert.That(calendarEvent.Start?.TimeZone, Is.EqualTo("UTC"));
            Assert.That(calendarEvent.Start?.DateTime, Is.EqualTo(start.UtcDateTime.ToString("o")));
            Assert.That(calendarEvent.Location, Is.Null);
            Assert.That(calendarEvent.Body, Is.Null);
        });
    }

    [Test]
    public void BuildEvent_maps_attendees_when_provided()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        var calendarEvent = OutlookGraphClient.BuildEvent(
            "Standup",
            start,
            start.AddHours(1),
            location: null,
            bodyText: null,
            attendeeAddresses: ["alice@example.com", "bob@example.com"]);

        Assert.That(
            calendarEvent.Attendees?.Select(attendee => attendee.EmailAddress?.Address),
            Is.EqualTo(new[] { "alice@example.com", "bob@example.com" }));
    }

    [Test]
    public void BuildEvent_sets_attendees_to_an_empty_not_null_list_when_an_empty_array_is_provided()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));

        var calendarEvent = OutlookGraphClient.BuildEvent(
            "Standup",
            start,
            start.AddHours(1),
            location: null,
            bodyText: null,
            attendeeAddresses: []);

        Assert.That(calendarEvent.Attendees, Is.Empty);
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_only_the_subject_when_only_subject_changes()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(toAddresses: null, subject: "New subject", bodyText: null);

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
        var message = OutlookGraphClient.BuildUpdateDraftMessage(toAddresses: null, subject: null, bodyText: "New body text.");

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.Null);
            Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Text));
            Assert.That(message.Body?.Content, Is.EqualTo("New body text."));
            Assert.That(message.ToRecipients, Is.Null);
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_only_the_recipients_when_only_recipients_change()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(
            toAddresses: ["alice@example.com", "bob@example.com"],
            subject: null,
            bodyText: null);

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.Null);
            Assert.That(message.Body, Is.Null);
            Assert.That(
                message.ToRecipients?.Select(recipient => recipient.EmailAddress?.Address),
                Is.EqualTo(new[] { "alice@example.com", "bob@example.com" }));
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_only_cc_when_only_cc_changes()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(
            toAddresses: null,
            subject: null,
            bodyText: null,
            ccAddresses: ["bob@example.com"]);

        Assert.Multiple(() =>
        {
            Assert.That(message.ToRecipients, Is.Null);
            Assert.That(message.BccRecipients, Is.Null);
            Assert.That(
                message.CcRecipients?.Select(recipient => recipient.EmailAddress?.Address),
                Is.EqualTo(new[] { "bob@example.com" }));
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_clears_cc_when_an_empty_array_is_provided()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(
            toAddresses: null,
            subject: null,
            bodyText: null,
            ccAddresses: []);

        Assert.That(message.CcRecipients, Is.Empty);
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_all_fields_when_all_are_provided()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(
            toAddresses: ["alice@example.com"],
            subject: "New subject",
            bodyText: "New body text.",
            ccAddresses: ["bob@example.com"],
            bccAddresses: ["carol@example.com"]);

        Assert.Multiple(() =>
        {
            Assert.That(message.Subject, Is.EqualTo("New subject"));
            Assert.That(message.Body?.ContentType, Is.EqualTo(BodyType.Text));
            Assert.That(message.Body?.Content, Is.EqualTo("New body text."));
            Assert.That(
                message.ToRecipients?.Select(recipient => recipient.EmailAddress?.Address),
                Is.EqualTo(new[] { "alice@example.com" }));
            Assert.That(
                message.CcRecipients?.Select(recipient => recipient.EmailAddress?.Address),
                Is.EqualTo(new[] { "bob@example.com" }));
            Assert.That(
                message.BccRecipients?.Select(recipient => recipient.EmailAddress?.Address),
                Is.EqualTo(new[] { "carol@example.com" }));
        });
    }

    [Test]
    public void BuildUpdateDraftMessage_sets_html_content_type_when_isHtml_is_true()
    {
        var message = OutlookGraphClient.BuildUpdateDraftMessage(
            toAddresses: null,
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
            location: null,
            bodyText: null,
            attendeeAddresses: null);

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
            location: null,
            bodyText: "New notes.",
            attendeeAddresses: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.Null);
            Assert.That(calendarEvent.Body?.ContentType, Is.EqualTo(BodyType.Text));
            Assert.That(calendarEvent.Body?.Content, Is.EqualTo("New notes."));
        });
    }

    [Test]
    public void BuildUpdateEvent_sets_only_start_and_end_when_only_the_time_changes()
    {
        var start = new DateTimeOffset(2026, 8, 1, 9, 0, 0, TimeSpan.FromHours(-5));
        var end = start.AddHours(1);

        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: start,
            end: end,
            location: null,
            bodyText: null,
            attendeeAddresses: null);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.Null);
            Assert.That(calendarEvent.Start?.DateTime, Is.EqualTo(start.UtcDateTime.ToString("o")));
            Assert.That(calendarEvent.End?.DateTime, Is.EqualTo(end.UtcDateTime.ToString("o")));
            Assert.That(calendarEvent.Location, Is.Null);
        });
    }

    [Test]
    public void BuildUpdateEvent_sets_only_the_location_when_only_location_changes()
    {
        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: null,
            end: null,
            location: "Conference Room B",
            bodyText: null,
            attendeeAddresses: null);

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
            location: null,
            bodyText: null,
            attendeeAddresses: ["alice@example.com"]);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.Null);
            Assert.That(
                calendarEvent.Attendees?.Select(attendee => attendee.EmailAddress?.Address),
                Is.EqualTo(new[] { "alice@example.com" }));
        });
    }

    [Test]
    public void BuildUpdateEvent_clears_attendees_to_an_empty_not_null_list_when_an_empty_array_is_provided()
    {
        var calendarEvent = OutlookGraphClient.BuildUpdateEvent(
            subject: null,
            start: null,
            end: null,
            location: null,
            bodyText: null,
            attendeeAddresses: []);

        Assert.That(calendarEvent.Attendees, Is.Empty);
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
            location: "Conference Room B",
            bodyText: "New notes.",
            attendeeAddresses: ["alice@example.com"]);

        Assert.Multiple(() =>
        {
            Assert.That(calendarEvent.Subject, Is.EqualTo("New subject"));
            Assert.That(calendarEvent.Start?.DateTime, Is.EqualTo(start.UtcDateTime.ToString("o")));
            Assert.That(calendarEvent.End?.DateTime, Is.EqualTo(end.UtcDateTime.ToString("o")));
            Assert.That(calendarEvent.Location?.DisplayName, Is.EqualTo("Conference Room B"));
            Assert.That(calendarEvent.Body?.Content, Is.EqualTo("New notes."));
            Assert.That(
                calendarEvent.Attendees?.Select(attendee => attendee.EmailAddress?.Address),
                Is.EqualTo(new[] { "alice@example.com" }));
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
}
