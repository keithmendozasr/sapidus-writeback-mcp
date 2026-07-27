using Microsoft.Graph.Models;

namespace OutlookWriteback.Graph.Confirmation;

/// <summary>
/// Orchestrates delete_event's two-step, confirmation-gated flow: the first call fetches the
/// event and issues a confirmation token without deleting anything; the second call only
/// deletes once that same token validates for the same event ID.
/// </summary>
public sealed class EventDeletionService(OutlookGraphClient client, DeleteConfirmationTokenService tokenService)
{
    public async Task<PendingEventDeletion> RequestDeletionAsync(string eventId, CancellationToken cancellationToken = default)
    {
        var calendarEvent = await client.GetEventByIdAsync(eventId, cancellationToken);
        var confirmationToken = tokenService.Issue(eventId);

        return new PendingEventDeletion(eventId, calendarEvent?.Subject, calendarEvent?.Start, calendarEvent?.End, confirmationToken);
    }

    public async Task<bool> ConfirmDeletionAsync(string eventId, string confirmationToken, CancellationToken cancellationToken = default)
    {
        if (!tokenService.Validate(eventId, confirmationToken))
            return false;

        await client.DeleteEventAsync(eventId, cancellationToken);

        return true;
    }
}

public sealed record PendingEventDeletion(
    string EventId,
    string? Subject,
    DateTimeTimeZone? Start,
    DateTimeTimeZone? End,
    string ConfirmationToken);
