using Microsoft.Graph.Models;
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
}
