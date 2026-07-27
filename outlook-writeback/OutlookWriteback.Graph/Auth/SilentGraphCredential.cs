using Azure.Core;

namespace OutlookWriteback.Graph.Auth;

/// <summary>
/// TokenCredential that silently redeems a cached refresh token instead of prompting a user -
/// the non-interactive counterpart to InteractiveBrowserCredential, for a deployed service
/// that can't pop a browser. Caches the last issued access token in memory with a safety
/// buffer so it doesn't rotate the refresh token on every single Graph call.
/// </summary>
public sealed class SilentGraphCredential(GraphTokenEndpointClient tokenEndpointClient, IRefreshTokenStore store, string[] scopes)
    : TokenCredential
{
    private static readonly TimeSpan RefreshBuffer = TimeSpan.FromMinutes(5);

    private readonly SemaphoreSlim _lock = new(1, 1);

    private AccessToken? _cachedToken;

    // requestContext.Scopes is ignored deliberately - this credential always redeems the
    // fixed Mail.ReadWrite/Calendars.ReadWrite scopes the app was registered for.
    public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken) =>
        GetTokenAsync(requestContext, cancellationToken).AsTask().GetAwaiter().GetResult();

    public override async ValueTask<AccessToken> GetTokenAsync(
        TokenRequestContext requestContext,
        CancellationToken cancellationToken)
    {
        if (_cachedToken is { } cached && cached.ExpiresOn > DateTimeOffset.UtcNow + RefreshBuffer)
            return cached;

        await _lock.WaitAsync(cancellationToken);

        try
        {
            if (_cachedToken is { } recheck && recheck.ExpiresOn > DateTimeOffset.UtcNow + RefreshBuffer)
                return recheck;

            var currentRefreshToken = await store.GetRefreshTokenAsync(cancellationToken);
            var response = await tokenEndpointClient.RedeemRefreshTokenAsync(currentRefreshToken, scopes, cancellationToken);

            if (response.RefreshToken is { } rotated)
                await store.SaveRefreshTokenAsync(rotated, cancellationToken);

            _cachedToken = new AccessToken(response.AccessToken, DateTimeOffset.UtcNow.AddSeconds(response.ExpiresInSeconds));

            return _cachedToken.Value;
        }
        finally
        {
            _lock.Release();
        }
    }
}
