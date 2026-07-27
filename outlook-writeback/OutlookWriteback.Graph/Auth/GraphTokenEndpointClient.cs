using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace OutlookWriteback.Graph.Auth;

public sealed record GraphTokenResponse(string AccessToken, string? RefreshToken, int ExpiresInSeconds);

/// <summary>
/// Talks to the Microsoft identity platform's v2.0 token endpoint directly instead of going
/// through MSAL - this app only ever redeems a single refresh token for a single user, so
/// MSAL's multi-account token cache machinery has nothing to add here.
/// </summary>
public sealed class GraphTokenEndpointClient
{
    private readonly HttpClient _httpClient;
    private readonly string _tenantId;
    private readonly string _clientId;

    public GraphTokenEndpointClient(string tenantId, string clientId, HttpClient? httpClient = null)
    {
        _tenantId = tenantId;
        _clientId = clientId;
        _httpClient = httpClient ?? new HttpClient();
    }

    public Task<GraphTokenResponse> RedeemRefreshTokenAsync(
        string refreshToken,
        IEnumerable<string> scopes,
        CancellationToken cancellationToken = default) =>
        PostTokenRequestAsync(
            new Dictionary<string, string>
            {
                ["grant_type"] = "refresh_token",
                ["refresh_token"] = refreshToken,
                ["client_id"] = _clientId,
                ["scope"] = string.Join(' ', scopes.Append("offline_access")),
            },
            cancellationToken);

    public Task<GraphTokenResponse> ExchangeAuthorizationCodeAsync(
        string code,
        string codeVerifier,
        string redirectUri,
        IEnumerable<string> scopes,
        CancellationToken cancellationToken = default) =>
        PostTokenRequestAsync(
            new Dictionary<string, string>
            {
                ["grant_type"] = "authorization_code",
                ["code"] = code,
                ["code_verifier"] = codeVerifier,
                ["redirect_uri"] = redirectUri,
                ["client_id"] = _clientId,
                ["scope"] = string.Join(' ', scopes.Append("offline_access")),
            },
            cancellationToken);

    private async Task<GraphTokenResponse> PostTokenRequestAsync(
        Dictionary<string, string> form,
        CancellationToken cancellationToken)
    {
        var response = await _httpClient.PostAsync(
            $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token",
            new FormUrlEncodedContent(form),
            cancellationToken);

        response.EnsureSuccessStatusCode();

        var payload = await response.Content.ReadFromJsonAsync<TokenEndpointPayload>(cancellationToken: cancellationToken);

        return new GraphTokenResponse(payload!.AccessToken, payload.RefreshToken, payload.ExpiresIn);
    }

    private sealed class TokenEndpointPayload
    {
        [JsonPropertyName("access_token")]
        public required string AccessToken { get; init; }

        [JsonPropertyName("refresh_token")]
        public string? RefreshToken { get; init; }

        [JsonPropertyName("expires_in")]
        public required int ExpiresIn { get; init; }
    }
}
