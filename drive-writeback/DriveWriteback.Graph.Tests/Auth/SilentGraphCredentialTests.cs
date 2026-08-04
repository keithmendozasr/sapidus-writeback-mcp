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

    [Test]
    public async Task GetTokenAsync_persists_the_rotated_refresh_token_returned_by_the_token_endpoint()
    {
        var store = new FakeRefreshTokenStore("initial-refresh-token");
        var handler = CreateHandler("new-access-token", "rotated-refresh-token");
        var credential = CreateCredential(handler, store);

        await credential.GetTokenAsync(new TokenRequestContext(["Files.ReadWrite.All"]), CancellationToken.None);

        Assert.That(await store.GetRefreshTokenAsync(), Is.EqualTo("rotated-refresh-token"));
    }

    [Test]
    public async Task GetTokenAsync_reuses_the_existing_refresh_token_when_the_response_omits_one()
    {
        var store = new FakeRefreshTokenStore("initial-refresh-token");
        var handler = CreateHandler("new-access-token", refreshToken: null);
        var credential = CreateCredential(handler, store);

        await credential.GetTokenAsync(new TokenRequestContext(["Files.ReadWrite.All"]), CancellationToken.None);

        var storedRefreshToken = await store.GetRefreshTokenAsync();

        Assert.Multiple(() =>
        {
            Assert.That(storedRefreshToken, Is.EqualTo("initial-refresh-token"));
            Assert.That(store.SaveCallCount, Is.EqualTo(0));
        });
    }

    [Test]
    public async Task GetTokenAsync_reuses_the_cached_access_token_without_redeeming_again_before_expiry()
    {
        var store = new FakeRefreshTokenStore("initial-refresh-token");
        var requestCount = 0;
        var handler = new StubHttpMessageHandler(_ =>
        {
            requestCount++;

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"access_token":"new-access-token","refresh_token":"rotated-refresh-token","expires_in":3600}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });
        var credential = CreateCredential(handler, store);

        await credential.GetTokenAsync(new TokenRequestContext(["Files.ReadWrite.All"]), CancellationToken.None);
        await credential.GetTokenAsync(new TokenRequestContext(["Files.ReadWrite.All"]), CancellationToken.None);

        Assert.That(requestCount, Is.EqualTo(1));
    }

    [Test]
    public async Task GetTokenAsync_redeems_again_once_the_cached_access_token_is_near_expiry()
    {
        var store = new FakeRefreshTokenStore("initial-refresh-token");
        var requestCount = 0;
        var handler = new StubHttpMessageHandler(_ =>
        {
            requestCount++;

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"access_token":"new-access-token","refresh_token":"rotated-refresh-token","expires_in":0}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });
        var credential = CreateCredential(handler, store);

        await credential.GetTokenAsync(new TokenRequestContext(["Files.ReadWrite.All"]), CancellationToken.None);
        await credential.GetTokenAsync(new TokenRequestContext(["Files.ReadWrite.All"]), CancellationToken.None);

        Assert.That(requestCount, Is.EqualTo(2));
    }
}
