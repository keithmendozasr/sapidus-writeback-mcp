using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("OutlookWriteback.Graph.Tests")]

namespace OutlookWriteback.Graph;

/// <summary>
/// Thin wrapper over the Graph SDK for the Phase 0 spike: proves delegated auth plus
/// POST /me/messages, POST /me/events, and GET /me/events/{id} work against a
/// self-registered Entra app. Not the full create_draft/create_event/etc. tool surface —
/// that lands in Phase 1-2 once this is verified.
/// </summary>
public sealed class OutlookGraphClient
{
    private readonly GraphServiceClient _client;

    public OutlookGraphClient(GraphServiceClient client)
    {
        _client = client;
    }

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

    public async Task<string?> CreateDraftAsync(
        string toAddress,
        string subject,
        string bodyText,
        CancellationToken cancellationToken = default)
    {
        var message = BuildDraftMessage(toAddress, subject, bodyText);
        var created = await _client.Me.Messages.PostAsync(message, cancellationToken: cancellationToken);

        return created?.Id;
    }

    public async Task<string?> CreateEventAsync(
        string subject,
        DateTimeOffset start,
        DateTimeOffset end,
        string? location = null,
        string? bodyText = null,
        CancellationToken cancellationToken = default)
    {
        var calendarEvent = BuildEvent(subject, start, end, location, bodyText);
        var created = await _client.Me.Events.PostAsync(calendarEvent, cancellationToken: cancellationToken);

        return created?.Id;
    }

    public Task<Event?> GetEventByIdAsync(string eventId, CancellationToken cancellationToken = default) =>
        _client.Me.Events[eventId].GetAsync(cancellationToken: cancellationToken);

    internal static Message BuildDraftMessage(string toAddress, string subject, string bodyText) => new()
    {
        Subject = subject,
        Body = new ItemBody { ContentType = BodyType.Text, Content = bodyText },
        ToRecipients = [new Recipient { EmailAddress = new EmailAddress { Address = toAddress } }],
    };

    internal static Event BuildEvent(
        string subject,
        DateTimeOffset start,
        DateTimeOffset end,
        string? location,
        string? bodyText)
    {
        var calendarEvent = new Event
        {
            Subject = subject,
            Start = ToGraphDateTime(start),
            End = ToGraphDateTime(end),
        };

        if (location is not null)
            calendarEvent.Location = new Location { DisplayName = location };

        if (bodyText is not null)
            calendarEvent.Body = new ItemBody { ContentType = BodyType.Text, Content = bodyText };

        return calendarEvent;
    }

    private static DateTimeTimeZone ToGraphDateTime(DateTimeOffset value) => new()
    {
        DateTime = value.UtcDateTime.ToString("o"),
        TimeZone = "UTC",
    };
}
