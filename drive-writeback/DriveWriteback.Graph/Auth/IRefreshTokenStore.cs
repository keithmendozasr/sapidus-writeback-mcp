namespace DriveWriteback.Graph.Auth;

/// <summary>
/// Persists the Graph refresh token for a non-interactive credential. Deliberately has no
/// Key Vault (or any other backing store) dependency here - that stays in whichever host
/// project implements this interface, so DriveWriteback.Graph never needs to change if the
/// storage backend changes.
/// </summary>
public interface IRefreshTokenStore
{
    Task<string> GetRefreshTokenAsync(CancellationToken cancellationToken = default);

    Task SaveRefreshTokenAsync(string refreshToken, CancellationToken cancellationToken = default);
}
