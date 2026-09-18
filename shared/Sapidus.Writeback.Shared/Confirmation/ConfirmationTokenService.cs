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

    // ASCII Unit Separator - a control character that can't appear in a Graph resource id,
    // used only to canonicalize a batch token's id set into the single opaque resourceId
    // string Issue/TryDecode already know how to carry. Never appears in the token string
    // itself (which is base64url text) - only in the decoded payload bytes.
    private const char IdSetDelimiter = '';

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
        if (!TryDecode(token, out var decodedResourceId, out var expiresAt))
            return false;

        return decodedResourceId == resourceId && _timeProvider.GetUtcNow() <= expiresAt;
    }

    /// <summary>
    /// Issues one token covering a whole set of resource ids, for a batch destructive action
    /// gated by a single confirmation round trip instead of one token per id. The set is
    /// canonicalized (deduplicated, sorted) before signing so token content doesn't depend on
    /// caller-supplied ordering or duplicates.
    /// </summary>
    public string IssueBatch(IEnumerable<string> resourceIds)
    {
        var canonicalIds = resourceIds.Distinct(StringComparer.Ordinal).OrderBy(id => id, StringComparer.Ordinal);

        return Issue(string.Join(IdSetDelimiter, canonicalIds));
    }

    /// <summary>
    /// Validates a batch token and reports which of the caller's requestedIds are authorized
    /// by it. TokenValid is false for a missing/tampered/expired/malformed token - an
    /// all-or-nothing failure, same as Validate. When TokenValid is true, AuthorizedIds is the
    /// subset of requestedIds that were part of the original IssueBatch call; an id outside
    /// the original set is simply absent from AuthorizedIds rather than invalidating the whole
    /// token - the caller is expected to reject that id individually.
    /// </summary>
    public BatchValidationResult ValidateSubset(IEnumerable<string> requestedIds, string token)
    {
        if (!TryDecode(token, out var encodedIdSet, out var expiresAt) || _timeProvider.GetUtcNow() > expiresAt)
            return new BatchValidationResult(TokenValid: false, AuthorizedIds: new HashSet<string>());

        var originalIds = encodedIdSet.Split(IdSetDelimiter).ToHashSet(StringComparer.Ordinal);
        var authorizedIds = requestedIds.Where(originalIds.Contains).ToHashSet(StringComparer.Ordinal);

        return new BatchValidationResult(TokenValid: true, AuthorizedIds: authorizedIds);
    }

    private bool TryDecode(string token, out string resourceId, out DateTimeOffset expiresAt)
    {
        resourceId = "";
        expiresAt = default;

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

        if (payloadParts.Length != 2 || !long.TryParse(payloadParts[1], out var expiresAtSeconds))
            return false;

        resourceId = payloadParts[0];
        expiresAt = DateTimeOffset.FromUnixTimeSeconds(expiresAtSeconds);

        return true;
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

/// <summary>
/// Result of ConfirmationTokenService.ValidateSubset. When TokenValid is false, AuthorizedIds
/// is always empty and the whole confirming call should be rejected (invalid/expired token).
/// When TokenValid is true, AuthorizedIds is the subset of the requested ids that were part of
/// the original batch - any requested id missing from it was not part of that batch and should
/// be rejected individually, not by failing the whole call.
/// </summary>
public sealed record BatchValidationResult(bool TokenValid, IReadOnlySet<string> AuthorizedIds);
