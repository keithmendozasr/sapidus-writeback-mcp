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
}
