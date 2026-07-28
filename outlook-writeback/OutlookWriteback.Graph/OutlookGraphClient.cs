using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using OutlookWriteback.Graph.Auth;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("OutlookWriteback.Graph.Tests")]

namespace OutlookWriteback.Graph;

/// <summary>
/// Thin wrapper over the Graph SDK for the Phase 0 spike: proves delegated auth plus
/// POST /me/messages, POST /me/events, and GET /me/events/{id} work against a
/// self-registered Entra app. Not the full create_draft/create_event/etc. tool surface —
/// that lands in Phase 1-2 once this is verified.
/// </summary>
public sealed class OutlookGraphClient(GraphServiceClient client)
{
    /// <summary>
    /// Requires the Entra app to be registered as a public client with
    /// "http://localhost" listed under Mobile and desktop redirect URIs - a
    /// confidential/web registration will fail the redirect-URI check at sign-in.
    /// </summary>
    public static OutlookGraphClient CreateWithInteractiveBrowserAuth(
        string tenantId,
        string clientId,
        IEnumerable<string> scopes)
    {
        var credential = new InteractiveBrowserCredential(
            new InteractiveBrowserCredentialOptions
            {
                TenantId = tenantId,
                ClientId = clientId,
                RedirectUri = new Uri("http://localhost"),
            });

        var client = new GraphServiceClient(credential, scopes);

        return new OutlookGraphClient(client);
    }

    /// <summary>
    /// Non-interactive counterpart to CreateWithInteractiveBrowserAuth, for a deployed service
    /// that can't pop a browser - silently redeems a refresh token cached in refreshTokenStore
    /// instead. The refresh token itself must be seeded once via a separate interactive
    /// bootstrap step (see OutlookWriteback.Bootstrap).
    /// </summary>
    public static OutlookGraphClient CreateWithSilentRefreshAuth(
        string tenantId,
        string clientId,
        IEnumerable<string> scopes,
        IRefreshTokenStore refreshTokenStore)
    {
        var scopeArray = scopes as string[] ?? [.. scopes];
        var tokenEndpointClient = new GraphTokenEndpointClient(tenantId, clientId);
        var credential = new SilentGraphCredential(tokenEndpointClient, refreshTokenStore, scopeArray);
        var client = new GraphServiceClient(credential, scopeArray);

        return new OutlookGraphClient(client);
    }

    public async Task<string?> CreateDraftAsync(
        string toAddress,
        string subject,
        string bodyText,
        bool isHtml = false,
        CancellationToken cancellationToken = default)
    {
        var message = BuildDraftMessage(toAddress, subject, bodyText, isHtml);
        var created = await client.Me.Messages.PostAsync(message, cancellationToken: cancellationToken);

        return created?.Id;
    }

    public async Task<string?> UpdateDraftAsync(
        string draftId,
        string? toAddress = null,
        string? subject = null,
        string? bodyText = null,
        bool isHtml = false,
        CancellationToken cancellationToken = default)
    {
        if (toAddress is null && subject is null && bodyText is null)
            throw new ArgumentException("At least one of toAddress, subject, or bodyText must be provided.");

        var message = BuildUpdateDraftMessage(toAddress, subject, bodyText, isHtml);
        var updated = await client.Me.Messages[draftId].PatchAsync(message, cancellationToken: cancellationToken);

        return updated?.Id ?? draftId;
    }

    public async Task<string?> CreateEventAsync(
        string subject,
        DateTimeOffset start,
        DateTimeOffset end,
        string timeZone,
        string? location = null,
        string? bodyText = null,
        IEnumerable<string>? attendeeAddresses = null,
        int? reminderMinutesBeforeStart = null,
        CancellationToken cancellationToken = default)
    {
        if (reminderMinutesBeforeStart is < 0)
            throw new ArgumentException("reminderMinutesBeforeStart must not be negative.");

        var calendarEvent = BuildEvent(subject, start, end, timeZone, location, bodyText, attendeeAddresses, reminderMinutesBeforeStart);
        var created = await client.Me.Events.PostAsync(calendarEvent, cancellationToken: cancellationToken);

        return created?.Id;
    }

    public Task<Event?> GetEventByIdAsync(string eventId, CancellationToken cancellationToken = default) =>
        client.Me.Events[eventId].GetAsync(cancellationToken: cancellationToken);

    public async Task<string?> UpdateEventAsync(
        string eventId,
        string? subject = null,
        DateTimeOffset? start = null,
        DateTimeOffset? end = null,
        string? timeZone = null,
        string? location = null,
        string? bodyText = null,
        IEnumerable<string>? attendeeAddresses = null,
        int? reminderMinutesBeforeStart = null,
        CancellationToken cancellationToken = default)
    {
        if (reminderMinutesBeforeStart is < 0)
            throw new ArgumentException("reminderMinutesBeforeStart must not be negative.");

        if (timeZone is not null && start is null && end is null)
            throw new ArgumentException("timeZone can only be provided together with start and/or end.");

        if (subject is null && start is null && end is null && location is null && bodyText is null
            && attendeeAddresses is null && reminderMinutesBeforeStart is null)
        {
            throw new ArgumentException(
                "At least one of subject, start, end, location, bodyText, attendeeAddresses, or reminderMinutesBeforeStart must be provided.");
        }

        var calendarEvent = BuildUpdateEvent(subject, start, end, timeZone, location, bodyText, attendeeAddresses, reminderMinutesBeforeStart);
        var updated = await client.Me.Events[eventId].PatchAsync(calendarEvent, cancellationToken: cancellationToken);

        return updated?.Id ?? eventId;
    }

    public Task DeleteEventAsync(string eventId, CancellationToken cancellationToken = default) =>
        client.Me.Events[eventId].DeleteAsync(cancellationToken: cancellationToken);

    internal static Message BuildDraftMessage(string toAddress, string subject, string bodyText, bool isHtml = false) => new()
    {
        Subject = subject,
        Body = new ItemBody { ContentType = isHtml ? BodyType.Html : BodyType.Text, Content = bodyText },
        ToRecipients = [new Recipient { EmailAddress = new EmailAddress { Address = toAddress } }],
    };

    internal static Message BuildUpdateDraftMessage(string? toAddress, string? subject, string? bodyText, bool isHtml = false)
    {
        var message = new Message();

        if (subject is not null)
            message.Subject = subject;

        if (bodyText is not null)
            message.Body = new ItemBody { ContentType = isHtml ? BodyType.Html : BodyType.Text, Content = bodyText };

        if (toAddress is not null)
            message.ToRecipients = [new Recipient { EmailAddress = new EmailAddress { Address = toAddress } }];

        return message;
    }

    internal static Event BuildEvent(
        string subject,
        DateTimeOffset start,
        DateTimeOffset end,
        string timeZone,
        string? location,
        string? bodyText,
        IEnumerable<string>? attendeeAddresses = null,
        int? reminderMinutesBeforeStart = null)
    {
        var calendarEvent = new Event
        {
            Subject = subject,
            Start = ToGraphDateTime(start, timeZone),
            End = ToGraphDateTime(end, timeZone),
        };

        if (location is not null)
            calendarEvent.Location = new Location { DisplayName = location };

        if (bodyText is not null)
            calendarEvent.Body = new ItemBody { ContentType = BodyType.Text, Content = bodyText };

        if (attendeeAddresses is not null)
            calendarEvent.Attendees = BuildAttendees(attendeeAddresses);

        if (reminderMinutesBeforeStart is not null)
        {
            calendarEvent.IsReminderOn = true;
            calendarEvent.ReminderMinutesBeforeStart = reminderMinutesBeforeStart;
        }

        return calendarEvent;
    }

    private static List<Attendee> BuildAttendees(IEnumerable<string> attendeeAddresses) =>
        [.. attendeeAddresses.Select(address => new Attendee
        {
            EmailAddress = new EmailAddress { Address = address },
            Type = AttendeeType.Required,
        })];

    internal static Event BuildUpdateEvent(
        string? subject,
        DateTimeOffset? start,
        DateTimeOffset? end,
        string? timeZone,
        string? location,
        string? bodyText,
        IEnumerable<string>? attendeeAddresses,
        int? reminderMinutesBeforeStart)
    {
        var calendarEvent = new Event();

        if (subject is not null)
            calendarEvent.Subject = subject;

        if (start is not null)
            calendarEvent.Start = ToGraphDateTime(start.Value, timeZone);

        if (end is not null)
            calendarEvent.End = ToGraphDateTime(end.Value, timeZone);

        if (location is not null)
            calendarEvent.Location = new Location { DisplayName = location };

        if (bodyText is not null)
            calendarEvent.Body = new ItemBody { ContentType = BodyType.Text, Content = bodyText };

        if (attendeeAddresses is not null)
            calendarEvent.Attendees = BuildAttendees(attendeeAddresses);

        if (reminderMinutesBeforeStart is not null)
        {
            calendarEvent.IsReminderOn = true;
            calendarEvent.ReminderMinutesBeforeStart = reminderMinutesBeforeStart;
        }

        return calendarEvent;
    }

    /// <summary>
    /// When <paramref name="timeZoneId"/> is null (only reachable from an update that changes
    /// start/end without specifying a timezone), preserves the legacy behavior byte-for-byte:
    /// TimeZone "UTC" and a DateTime string carrying a trailing "Z". When a timezone id is given
    /// (always true for event creation, since it's required there), converts to that zone's local
    /// wall-clock time instead, with no offset/Z suffix - the format Graph expects for a named
    /// zone. Do not "fix" this into one consistent format: doing so would change the wire format
    /// for existing update callers that never pass a timezone.
    /// </summary>
    private static DateTimeTimeZone ToGraphDateTime(DateTimeOffset value, string? timeZoneId)
    {
        if (timeZoneId is null)
        {
            return new DateTimeTimeZone
            {
                DateTime = value.UtcDateTime.ToString("o"),
                TimeZone = "UTC",
            };
        }

        var timeZone = TimeZoneInfo.FindSystemTimeZoneById(timeZoneId);
        var localDateTime = TimeZoneInfo.ConvertTime(value, timeZone).DateTime;

        return new DateTimeTimeZone
        {
            DateTime = localDateTime.ToString("o"),
            TimeZone = timeZoneId,
        };
    }
}
