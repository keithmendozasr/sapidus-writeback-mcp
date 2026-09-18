using System.Collections;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using Microsoft.Extensions.Logging;
using OutlookWriteback.Graph.Confirmation;

namespace OutlookWriteback.Functions;

public sealed class DeleteEventTool(EventDeletionService deletionService, ILogger<DeleteEventTool> logger)
{
    [Function(nameof(DeleteEventTool))]
    public async Task<string> RunAsync(
        [McpToolTrigger(
            "delete_event",
            "Delete one or more calendar events (1-25 IDs). Two-step and confirmation-gated: call once with just " +
                "eventIds to preview the events and get a confirmationToken - this does NOT delete anything. Call " +
                "again with that confirmationToken and the eventIds you want deleted (any non-empty subset of the " +
                "originally previewed IDs - drop any you don't want deleted) to actually delete them. A single " +
                "delete is just a 1-element eventIds array. " +
                "If either call fails with an authentication/401-style error, tell the user the outlook-writeback " +
                "connector may need to be reconnected (Settings/Customize > Connectors > outlook-writeback > " +
                "Reconnect) before retrying - don't silently retry or fail.")]
            ToolInvocationContext context,
        [McpToolProperty("eventIds", "The calendar event IDs to delete (1-25 items; a single delete is a 1-element array).", isRequired: true)]
            string[]? eventIds,
        [McpToolProperty(
            "confirmationToken",
            "Omit on the first call. Supply the token returned by the first call to confirm - eventIds may then be " +
                "any non-empty subset of what was originally previewed.")]
            string? confirmationToken)
    {
        if (eventIds is null or [])
            return "At least one eventId is required.";

        if (eventIds.Length > EventDeletionService.MaxBatchSize)
            return $"Too many event IDs: {eventIds.Length} given, {EventDeletionService.MaxBatchSize} max per call. " +
                "Split this into multiple delete_event calls.";

        if (confirmationToken is null)
        {
            LogWithEventIds(logger, $"delete_event: outcome=preview requested-count={eventIds.Length}", eventIds);

            var preview = await deletionService.RequestDeletionAsync(eventIds);
            var lines = preview.Events.Select(e => e.Found
                ? $"- \"{e.Subject}\" ({FormatWhen(e)}, event ID: {e.EventId})"
                : $"- NOT FOUND: {e.EventId}");

            return "About to delete:\n" + string.Join("\n", lines) +
                "\n\nThis has NOT been deleted yet. If the user confirms, call delete_event again with the eventIds " +
                $"to delete (drop any you don't want deleted) and confirmationToken=\"{preview.ConfirmationToken}\".";
        }

        LogWithEventIds(logger, $"delete_event: outcome=confirmed requested-count={eventIds.Length}", eventIds);

        var result = await deletionService.ConfirmDeletionAsync(eventIds, confirmationToken);

        if (!result.TokenValid)
            return "That confirmation token is invalid or expired. Call delete_event again with just eventIds " +
                "(no confirmationToken) to get a new one.";

        var deletedCount = result.Outcomes.Count(o => o.Deleted);
        var failed = result.Outcomes.Where(o => !o.Deleted).ToList();
        var summary = $"{deletedCount} of {result.Outcomes.Count} deleted.";

        if (failed.Count > 0)
            summary += " Failed: " + string.Join(", ", failed.Select(f => $"{f.EventId} ({f.FailureReason})"));

        return summary;
    }

    private static string FormatWhen(PreviewedEventDeletion e) =>
        e.Start is null ? "unknown time" : $"{e.Start.DateTime} to {e.End?.DateTime} ({e.Start.TimeZone})";

    // Carries eventIds as structured state (an ApplicationInsights customDimension) instead of interpolating them into message.
    private static void LogWithEventIds(ILogger logger, string message, string[] eventIds) =>
        logger.Log(LogLevel.Information, eventId: default, new EventIdsLogState(message, eventIds), exception: null, static (state, _) => state.Message);

    private sealed class EventIdsLogState(string message, string[] eventIds) : IReadOnlyList<KeyValuePair<string, object?>>
    {
        public string Message { get; } = message;

        public int Count => 2;

        public KeyValuePair<string, object?> this[int index] => index switch
        {
            0 => new KeyValuePair<string, object?>("EventIds", eventIds),
            1 => new KeyValuePair<string, object?>("RequestedCount", eventIds.Length),
            _ => throw new ArgumentOutOfRangeException(nameof(index)),
        };

        public IEnumerator<KeyValuePair<string, object?>> GetEnumerator()
        {
            yield return this[0];
            yield return this[1];
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
