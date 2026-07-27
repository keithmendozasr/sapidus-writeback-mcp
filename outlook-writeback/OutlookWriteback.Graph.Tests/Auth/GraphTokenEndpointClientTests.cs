using System.Net;
using System.Text;
using OutlookWriteback.Graph.Auth;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests.Auth;

[TestFixture]
[Category("Unit")]
public class GraphTokenEndpointClientTests
{
    private static GraphTokenEndpointClient CreateClient(StubHttpMessageHandler handler) =>
        new("tenant-id", "client-id", new HttpClient(handler));

    [Test]
    public async Task RedeemRefreshTokenAsync_posts_grant_type_refresh_token_and_returns_parsed_tokens()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            var form = request.Content!.ReadAsStringAsync().GetAwaiter().GetResult();

            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Post));
                Assert.That(request.RequestUri!.ToString(), Does.Contain("/tenant-id/oauth2/v2.0/token"));
                Assert.That(form, Does.Contain("grant_type=refresh_token"));
                Assert.That(form, Does.Contain("refresh_token=old-refresh-token"));
            });

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"access_token":"new-access-token","refresh_token":"new-refresh-token","expires_in":3600}""",
                    Encoding.UTF8,
                    "application/json"),
            };
        });

        var response = await CreateClient(handler).RedeemRefreshTokenAsync(
            "old-refresh-token",
            ["Mail.ReadWrite", "Calendars.ReadWrite"]);

        Assert.Multiple(() =>
        {
            Assert.That(response.AccessToken, Is.EqualTo("new-access-token"));
            Assert.That(response.RefreshToken, Is.EqualTo("new-refresh-token"));
            Assert.That(response.ExpiresInSeconds, Is.EqualTo(3600));
        });
    }

    [Test]
    public async Task RedeemRefreshTokenAsync_includes_offline_access_in_the_requested_scope()
    {
        string? capturedForm = null;
        var handler = new StubHttpMessageHandler(request =>
        {
            capturedForm = request.Content!.ReadAsStringAsync().GetAwaiter().GetResult();

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"access_token":"new-access-token","refresh_token":"new-refresh-token","expires_in":3600}""",
                    Encoding.UTF8,
                    "application/json"),
            };
        });

        await CreateClient(handler).RedeemRefreshTokenAsync("old-refresh-token", ["Mail.ReadWrite"]);

        Assert.That(capturedForm, Does.Contain("scope=Mail.ReadWrite+offline_access"));
    }

    [Test]
    public void RedeemRefreshTokenAsync_throws_when_the_token_endpoint_returns_an_error_status()
    {
        var handler = new StubHttpMessageHandler(_ => new HttpResponseMessage(HttpStatusCode.BadRequest)
        {
            Content = new StringContent(
                """{"error":"invalid_grant","error_description":"Refresh token has expired."}""",
                Encoding.UTF8,
                "application/json"),
        });

        Assert.That(
            () => CreateClient(handler).RedeemRefreshTokenAsync("expired-refresh-token", ["Mail.ReadWrite"]),
            Throws.InstanceOf<HttpRequestException>());
    }

    [Test]
    public async Task ExchangeAuthorizationCodeAsync_posts_grant_type_authorization_code_with_the_code_verifier()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            var form = request.Content!.ReadAsStringAsync().GetAwaiter().GetResult();

            Assert.Multiple(() =>
            {
                Assert.That(form, Does.Contain("grant_type=authorization_code"));
                Assert.That(form, Does.Contain("code=auth-code"));
                Assert.That(form, Does.Contain("code_verifier=verifier-value"));
                Assert.That(form, Does.Contain("redirect_uri=http"));
            });

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"access_token":"first-access-token","refresh_token":"first-refresh-token","expires_in":3600}""",
                    Encoding.UTF8,
                    "application/json"),
            };
        });

        var response = await CreateClient(handler).ExchangeAuthorizationCodeAsync(
            "auth-code",
            "verifier-value",
            "http://localhost",
            ["Mail.ReadWrite", "Calendars.ReadWrite"]);

        Assert.Multiple(() =>
        {
            Assert.That(response.AccessToken, Is.EqualTo("first-access-token"));
            Assert.That(response.RefreshToken, Is.EqualTo("first-refresh-token"));
        });
    }
}
