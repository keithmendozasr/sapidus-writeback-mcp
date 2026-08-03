using System.Net;
using System.Text;
using Azure.Core;
using DriveWriteback.Graph.Auth;
using DriveWriteback.Graph.Tests.TestSupport;

namespace DriveWriteback.Graph.Tests.Auth;

[TestFixture]
[Category("Unit")]
public class SilentGraphCredentialTests
{
    private static SilentGraphCredential CreateCredential(StubHttpMessageHandler handler, FakeRefreshTokenStore store)
    {
        var tokenEndpointClient = new GraphTokenEndpointClient("tenant-id", "client-id", new HttpClient(handler));

        return new SilentGraphCredential(tokenEndpointClient, store, ["Files.ReadWrite.All"]);
    }

    private static StubHttpMessageHandler CreateHandler(string accessToken, string? refreshToken, int expiresIn = 3600) =>
        new(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(
                $$"""{"access_token":"{{accessToken}}","refresh_token":{{(refreshToken is null ? "null" : $"\"{refreshToken}\"")}},"expires_in":{{expiresIn}}}""",
                Encoding.UTF8,
                "application/json"),
        }));

    [Test]
    public async Task GetTokenAsync_redeems_the_stored_refresh_token_and_returns_an_access_token()
    {
        var store = new FakeRefreshTokenStore("initial-refresh-token");
        var handler = CreateHandler("new-access-token", "rotated-refresh-token");
        var credential = CreateCredential(handler, store);

        var token = await credential.GetTokenAsync(new TokenRequestContext(["Files.ReadWrite.All"]), CancellationToken.None);

        Assert.That(token.Token, Is.EqualTo("new-access-token"));
    }
}
