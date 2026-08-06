using System.Security.Cryptography;
using System.Text;

namespace Sapidus.Writeback.Shared.Confirmation;

/// <summary>
/// Issues and validates the confirmation token that gates a destructive tool's second,
/// actually-mutating call - the two-step pattern first established by outlook-writeback's
/// delete_event and reused by any server's own destructive-action confirmation flow (e.g.
/// drive-writeback's delete_item). Deliberately stateless - no Table Storage/Redis for the
/// "few minutes' TTL" a caller needs: the token itself carries the resource id and an expiry,
/// HMAC-signed so a client cannot mint or alter one. That unforgeability is what forces the
/// real two-step round-trip - an LLM client can't just compute a valid token and one-shot the
/// destructive call.
///
/// This is generic over an opaque resource id string and carries no Graph dependency or
/// server identity - it doesn't know or care whether that id names an email event, a drive
/// item, or anything else. That's what makes it shared/-eligible: unlike boundary-7a auth
/// code (deliberately kept server-local per REPO-CONVENTIONS.md §5, since sharing Graph-auth
/// mechanics would couple servers' Graph-scope isolation), this code never touches Graph or
/// credentials, so that coupling risk doesn't apply. Each server still passes in its own
/// signing key from its own Key Vault secret - only the mechanism is shared, not the key
/// material or the runtime state.
/// </summary>
public sealed class ConfirmationTokenService(byte[] signingKey, TimeProvider? timeProvider = null)
{
    private static readonly TimeSpan Ttl = TimeSpan.FromMinutes(5);

    private readonly TimeProvider _timeProvider = timeProvider ?? TimeProvider.System;

    public string Issue(string resourceId)
    {
        var expiresAt = _timeProvider.GetUtcNow().Add(Ttl).ToUnixTimeSeconds();
        var payload = Encoding.UTF8.GetBytes($"{resourceId}|{expiresAt}");
        var signature = Sign(payload);

        return $"{Base64UrlEncode(payload)}.{Base64UrlEncode(signature)}";
    }

    public bool Validate(string resourceId, string token)
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

        if (payloadParts.Length != 2 || payloadParts[0] != resourceId || !long.TryParse(payloadParts[1], out var expiresAt))
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
