using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests;

/// <summary>
/// Fast, offline tier: exercises OutlookGraphClient's real request-building and response
/// deserialization through the Graph SDK, against a stubbed HTTP handler instead of the
/// network. No credentials, no live tenant, safe for CI. For real checks against the actual
/// Graph API, see OutlookGraphClientE2ETests.
/// </summary>
[TestFixture]
[Category("Integration")]
public class OutlookGraphClientIntegrationTests
{
    private static OutlookGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new OutlookGraphClient(graphClient);
    }

    [Test]
    public async Task CreateDraftAsync_sends_a_POST_to_me_messages_and_returns_the_created_id()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Post));
                Assert.That(request.RequestUri!.AbsolutePath, Does.Contain("/me/messages"));
            });

            return new HttpResponseMessage(HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAMk-fake-draft-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        var draftId = await CreateClient(handler).CreateDraftAsync("owner@example.com", "Subject", "Body");

        Assert.That(draftId, Is.EqualTo("AAMk-fake-draft-id"));
    }
}
