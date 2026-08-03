using Azure.Security.KeyVault.Secrets;
using DriveWriteback.Graph.Auth;

namespace DriveWriteback.Auth;

/// <summary>
/// Reads and writes the cached Graph refresh token via Key Vault's Secrets SDK directly,
/// not a Key Vault-reference app setting - the app must be able to write the rotated token
/// back after every redemption, and app-setting references are read-only.
/// </summary>
public sealed class KeyVaultRefreshTokenStore(SecretClient secretClient) : IRefreshTokenStore
{
    private const string SecretName = "graph-refresh-token";

    public async Task<string> GetRefreshTokenAsync(CancellationToken cancellationToken = default)
    {
        var secret = await secretClient.GetSecretAsync(SecretName, cancellationToken: cancellationToken);

        return secret.Value.Value;
    }

    public async Task SaveRefreshTokenAsync(string refreshToken, CancellationToken cancellationToken = default) =>
        await secretClient.SetSecretAsync(SecretName, refreshToken, cancellationToken);
}
