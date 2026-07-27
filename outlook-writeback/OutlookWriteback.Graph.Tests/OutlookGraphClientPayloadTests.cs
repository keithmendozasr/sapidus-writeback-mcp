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
    public void UpdateDraftAsync_throws_when_no_fields_are_provided()
    {
        var httpClient = new HttpClient { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var client = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));

        Assert.That(
            () => client.UpdateDraftAsync("AAMk-fake-draft-id"),
            Throws.ArgumentException);
    }
}
