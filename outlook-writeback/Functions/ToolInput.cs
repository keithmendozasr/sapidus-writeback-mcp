using System.Globalization;

namespace OutlookWriteback.Functions;

/// <summary>
/// Helpers for turning bad caller input into something the calling agent can act on. An exception thrown out
/// of a tool surfaces to the agent only as a generic failure (the message stays in the server log), so input
/// problems are caught and returned as the tool's normal text result instead - the same approach delete_event
/// takes for its own validation.
/// </summary>
internal static class ToolInput
{
    // ISO 8601 date-time with an explicit offset (Z or +/-hh:mm), optional seconds / fractional seconds. A lenient
    // DateTimeOffset.Parse would also accept things like "10am" or an offset-less time and silently assume the
    // server's date and zone, which puts the event at the wrong time.
    private static readonly string[] OffsetFormats =
    [
        "yyyy-MM-dd'T'HH:mmzzz",
        "yyyy-MM-dd'T'HH:mm:sszzz",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz",
    ];

    private static readonly string[] UtcFormats =
    [
        "yyyy-MM-dd'T'HH:mm'Z'",
        "yyyy-MM-dd'T'HH:mm:ss'Z'",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'",
    ];

    public static DateTimeOffset ParseTimestamp(string field, string value)
    {
        var text = value?.Trim();

        if (DateTimeOffset.TryParseExact(text, OffsetFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            || DateTimeOffset.TryParseExact(text, UtcFormats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out parsed))
        {
            return parsed;
        }

        throw new ArgumentException(
            $"{field} must be an ISO 8601 date-time with a UTC offset, e.g. 2026-08-01T09:00:00-05:00 (got '{value}').");
    }

    /// <param name="outcome">What the caller can rely on, e.g. "No event was created."</param>
    public static string Rejected(string outcome, ArgumentException ex) =>
        $"{outcome} The input was rejected: {ex.Message} Correct it and call the tool again.";
}
