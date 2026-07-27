using Azure.Security.KeyVault.Secrets;
using OutlookWriteback.Graph.Auth;

namespace OutlookWriteback.Auth;

/// <summary>
/// Reads and writes the cached Graph refresh token via Key Vault's Secrets SDK directly,
/// not a Key Vault-reference app setting - the app must be able to write the rotated token
/// back after every redemption, and app-setting references are read-only.
/// </summary>
public sealed class KeyVaultRefreshTokenStore : IRefreshTokenStore
{
    private const string SecretName = "graph-refresh-token";

    private readonly SecretClient _secretClient;

    public KeyVaultRefreshTokenStore(SecretClient secretClient)
    {
        _secretClient = secretClient;
    }

    public async Task<string> GetRefreshTokenAsync(CancellationToken cancellationToken = default)
    {
        var secret = await _secretClient.GetSecretAsync(SecretName, cancellationToken: cancellationToken);

        return secret.Value.Value;
    }

    public async Task SaveRefreshTokenAsync(string refreshToken, CancellationToken cancellationToken = default) =>
        await _secretClient.SetSecretAsync(SecretName, refreshToken, cancellationToken);
}
