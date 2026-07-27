using System.Security.Cryptography;
using System.Text;

namespace OutlookWriteback.Graph.Confirmation;

/// <summary>
/// Issues and validates the confirmation token that gates delete_event's second, destructive
/// call. Deliberately stateless - no Table Storage/Redis for the "few minutes' TTL" the PRD
/// calls for: the token itself carries the event ID and an expiry, HMAC-signed so a client
/// cannot mint or alter one. That unforgeability is what forces the real two-step round-trip -
/// an LLM client can't just compute a valid token and one-shot the delete.
/// </summary>
public sealed class DeleteConfirmationTokenService(byte[] signingKey, TimeProvider? timeProvider = null)
{
    private static readonly TimeSpan Ttl = TimeSpan.FromMinutes(5);

    private readonly TimeProvider _timeProvider = timeProvider ?? TimeProvider.System;

    public string Issue(string eventId)
    {
        var expiresAt = _timeProvider.GetUtcNow().Add(Ttl).ToUnixTimeSeconds();
        var payload = Encoding.UTF8.GetBytes($"{eventId}|{expiresAt}");
        var signature = Sign(payload);

        return $"{Base64UrlEncode(payload)}.{Base64UrlEncode(signature)}";
    }

    public bool Validate(string eventId, string token)
    {
        var parts = token.Split('.');

        if (parts.Length != 2)
            return false;

        byte[] payload, signature;

        try
        {
            payload = Base64UrlDecode(parts[0]);
            signature = Base64UrlDecode(parts[1]);
        }
        catch (FormatException)
        {
            return false;
        }

        if (!CryptographicOperations.FixedTimeEquals(signature, Sign(payload)))
            return false;

        var payloadParts = Encoding.UTF8.GetString(payload).Split('|', 2);

        if (payloadParts.Length != 2 || payloadParts[0] != eventId || !long.TryParse(payloadParts[1], out var expiresAt))
            return false;

        return _timeProvider.GetUtcNow() <= DateTimeOffset.FromUnixTimeSeconds(expiresAt);
    }

    private byte[] Sign(byte[] payload) => HMACSHA256.HashData(signingKey, payload);

    private static string Base64UrlEncode(byte[] bytes) =>
        Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    private static byte[] Base64UrlDecode(string value)
    {
        var padded = value.Replace('-', '+').Replace('_', '/');
        padded = padded.PadRight(padded.Length + ((4 - (padded.Length % 4)) % 4), '=');

        return Convert.FromBase64String(padded);
    }
}
