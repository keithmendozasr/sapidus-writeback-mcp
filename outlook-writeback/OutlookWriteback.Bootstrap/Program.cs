using System.Net;
using System.Security.Cryptography;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using OutlookWriteback.Graph.Auth;

const int RedirectPort = 8400;
var redirectUri = $"http://localhost:{RedirectPort}/";
const string SecretName = "graph-refresh-token";
var scopes = new[] { "Mail.ReadWrite", "Calendars.ReadWrite" };

var tenantId = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_TENANT_ID")
    ?? throw new InvalidOperationException("OUTLOOK_WRITEBACK_TENANT_ID is not set.");

var clientId = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_CLIENT_ID")
    ?? throw new InvalidOperationException("OUTLOOK_WRITEBACK_CLIENT_ID is not set.");

var keyVaultUri = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_KEY_VAULT_URI")
    ?? throw new InvalidOperationException("OUTLOOK_WRITEBACK_KEY_VAULT_URI is not set.");

var (codeVerifier, codeChallenge) = GeneratePkcePair();

using var listener = new HttpListener();
listener.Prefixes.Add(redirectUri);
listener.Start();

var authorizeUrl =
    $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/authorize" +
    $"?client_id={Uri.EscapeDataString(clientId)}" +
    "&response_type=code" +
    $"&redirect_uri={Uri.EscapeDataString(redirectUri)}" +
    $"&scope={Uri.EscapeDataString(string.Join(' ', scopes.Append("offline_access")))}" +
    $"&code_challenge={codeChallenge}" +
    "&code_challenge_method=S256";

Console.WriteLine("Opening your browser to sign in to your Microsoft 365 tenant...");
Console.WriteLine("If it doesn't open automatically, visit this URL:");
Console.WriteLine(authorizeUrl);
System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(authorizeUrl) { UseShellExecute = true });

var context = await listener.GetContextAsync();
var code = context.Request.QueryString["code"];
var error = context.Request.QueryString["error"];

var responseBody = error is null
    ? "Sign-in complete. You can close this tab and return to the console."u8.ToArray()
    : "Sign-in failed. You can close this tab and return to the console."u8.ToArray();
context.Response.ContentType = "text/plain";
context.Response.OutputStream.Write(responseBody);
context.Response.OutputStream.Close();
listener.Stop();

if (error is not null || code is null)
    throw new InvalidOperationException($"Sign-in failed: {error ?? "no authorization code returned"}.");

var tokenEndpointClient = new GraphTokenEndpointClient(tenantId, clientId);
var tokenResponse = await tokenEndpointClient.ExchangeAuthorizationCodeAsync(code, codeVerifier, redirectUri, scopes);

if (tokenResponse.RefreshToken is not { } refreshToken)
    throw new InvalidOperationException("Token endpoint did not return a refresh token - check that offline_access was granted.");

var secretClient = new SecretClient(new Uri(keyVaultUri), new DefaultAzureCredential());
await secretClient.SetSecretAsync(SecretName, refreshToken);

Console.WriteLine("Refresh token written to Key Vault. The deployed app can now redeem it silently.");

static (string Verifier, string Challenge) GeneratePkcePair()
{
    var verifierBytes = RandomNumberGenerator.GetBytes(32);
    var verifier = Base64UrlEncode(verifierBytes);
    var challengeBytes = SHA256.HashData(System.Text.Encoding.ASCII.GetBytes(verifier));
    var challenge = Base64UrlEncode(challengeBytes);

    return (verifier, challenge);
}

static string Base64UrlEncode(byte[] bytes) =>
    Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
