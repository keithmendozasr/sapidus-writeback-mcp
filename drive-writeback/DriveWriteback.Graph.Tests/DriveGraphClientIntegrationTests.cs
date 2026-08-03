using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using DriveWriteback.Graph.Tests.TestSupport;

namespace DriveWriteback.Graph.Tests;

/// <summary>
/// Fast, offline tier: exercises DriveGraphClient's real request-building and response
/// deserialization through the Graph SDK, against a stubbed HTTP handler instead of the
/// network. No credentials, no live tenant, safe for CI. For real checks against the
/// actual Graph API, see DriveGraphClientE2ETests.
/// </summary>
[TestFixture]
[Category("Integration")]
public class DriveGraphClientIntegrationTests
{
    private static DriveGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new DriveGraphClient(graphClient);
    }

    [Test]
    public async Task TryGetItemByPathAsync_returns_null_on_a_stubbed_404()
    {
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
        {
            Content = new StringContent(
                """{"error":{"code":"itemNotFound","message":"Item not found"}}""",
                Encoding.UTF8,
                "application/json"),
        }));

        var result = await CreateClient(handler).TryGetItemByPathAsync("does-not-exist.md", driveId: "drive-id");

        Assert.That(result, Is.Null);
    }

    [Test]
    public async Task ResolveFolderPathAsync_stops_at_the_first_missing_segment()
    {
        // "a" resolves to an existing folder; any other segment name ("b") is reported
        // as not found, so the walk should stop there and report "b" and "c" as missing.
        var handler = new StubHttpMessageHandler(request =>
        {
            var respondsWithA = request.RequestUri!.ToString().Contains("%27a%27");

            var body = respondsWithA
                ? """{"value":[{"id":"a-id","name":"a","folder":{}}]}"""
                : """{"value":[]}""";

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json"),
            });
        });

        var resolution = await CreateClient(handler).ResolveFolderPathAsync("a/b/c", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(resolution.DeepestExisting?.Id, Is.EqualTo("a-id"));
            Assert.That(resolution.MissingSegments, Is.EqualTo(new[] { "b", "c" }));
        });
    }
}
