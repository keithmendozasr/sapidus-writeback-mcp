using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Sapidus.Writeback.Shared.Confirmation;

namespace OutlookWriteback.Graph.Confirmation;

/// <summary>
/// Orchestrates delete_event's two-step, confirmation-gated flow for a batch of events: the
/// first call fetches each event (best-effort) and issues one confirmation token covering the
/// whole requested set without deleting anything; the second call deletes only the subset of
/// that set the caller resends, validated against the same token.
/// </summary>
public sealed class EventDeletionService(OutlookGraphClient client, ConfirmationTokenService tokenService)
{
    /// <summary>
    /// Bounds blast radius and keeps the preview response readable - not a Graph-imposed limit.
    /// Enforced here as a defense-in-depth guard for any direct caller of this service; the
    /// MCP tool layer (DeleteEventTool) enforces it first and returns friendly, limit-naming
    /// text, since a thrown exception's message does not reach the caller through the Azure
    /// Functions MCP extension.
    /// </summary>
    public const int MaxBatchSize = 25;

    public async Task<BatchDeletionPreview> RequestDeletionAsync(
        IReadOnlyCollection<string> eventIds, CancellationToken cancellationToken = default)
    {
        if (eventIds.Count > MaxBatchSize)
            throw new ArgumentException($"Too many event IDs: {eventIds.Count} given, {MaxBatchSize} max per call.", nameof(eventIds));

        var canonicalIds = eventIds.Distinct(StringComparer.Ordinal).OrderBy(id => id, StringComparer.Ordinal).ToArray();
        var confirmationToken = tokenService.IssueBatch(canonicalIds);

        var previews = new List<PreviewedEventDeletion>();

        foreach (var eventId in canonicalIds)
        {
            try
            {
                var calendarEvent = await client.GetEventByIdAsync(eventId, cancellationToken);
                previews.Add(new PreviewedEventDeletion(eventId, calendarEvent?.Subject, calendarEvent?.Start, calendarEvent?.End, Found: true));
            }
            catch (ApiException)
            {
                previews.Add(new PreviewedEventDeletion(eventId, null, null, null, Found: false));
            }
        }

        return new BatchDeletionPreview(previews, confirmationToken);
    }

    public async Task<BatchDeletionResult> ConfirmDeletionAsync(
        IReadOnlyCollection<string> eventIds, string confirmationToken, CancellationToken cancellationToken = default)
    {
        var validation = tokenService.ValidateSubset(eventIds, confirmationToken);

        if (!validation.TokenValid)
            return new BatchDeletionResult([], TokenValid: false);

        var outcomes = new List<EventDeletionOutcome>();

        foreach (var eventId in eventIds.Distinct(StringComparer.Ordinal))
        {
            if (!validation.AuthorizedIds.Contains(eventId))
            {
                outcomes.Add(new EventDeletionOutcome(eventId, Deleted: false, "not part of the previewed set"));
                continue;
            }

            try
            {
                await client.DeleteEventAsync(eventId, cancellationToken);
                outcomes.Add(new EventDeletionOutcome(eventId, Deleted: true, FailureReason: null));
            }
            catch (ApiException ex)
            {
                outcomes.Add(new EventDeletionOutcome(eventId, Deleted: false, ex.Message));
            }
        }

        return new BatchDeletionResult(outcomes, TokenValid: true);
    }
}

public sealed record PreviewedEventDeletion(
    string EventId,
    string? Subject,
    DateTimeTimeZone? Start,
    DateTimeTimeZone? End,
    bool Found);

public sealed record BatchDeletionPreview(IReadOnlyList<PreviewedEventDeletion> Events, string ConfirmationToken);

public sealed record EventDeletionOutcome(string EventId, bool Deleted, string? FailureReason);

public sealed record BatchDeletionResult(IReadOnlyList<EventDeletionOutcome> Outcomes, bool TokenValid);
