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

    [Test]
    public async Task CreateFileAsync_replace_sends_a_single_unconditional_PUT_with_no_pre_check()
    {
        var requestCount = 0;
        var handler = new StubHttpMessageHandler(request =>
        {
            requestCount++;

            Assert.That(request.Method, Is.EqualTo(HttpMethod.Put));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"new-id"}""", Encoding.UTF8, "application/json"),
            });
        });

        var result = await CreateClient(handler).CreateFileAsync("notes.md", "content", conflictBehavior: "replace", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result?.Id, Is.EqualTo("new-id"));
            Assert.That(requestCount, Is.EqualTo(1), "replace should not send any pre-check request before the PUT.");
        });
    }

    [Test]
    public void CreateFileAsync_fail_pre_checks_with_a_GET_and_throws_without_ever_sending_a_PUT()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            // "fail" must never PUT once the pre-check GET reports the target already
            // exists - a PUT here would mean the client-side check was bypassed.
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Get));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"existing-id"}""", Encoding.UTF8, "application/json"),
            });
        });

        Assert.That(
            () => CreateClient(handler).CreateFileAsync("notes.md", "content", conflictBehavior: "fail", driveId: "drive-id"),
            Throws.InstanceOf<DriveItemAlreadyExistsException>());
    }

    [Test]
    public void CreateFileAsync_throws_DriveParentNotFoundException_when_the_parent_path_is_missing()
    {
        // Every request in this scenario is the parent-path existence check, which
        // reports "not found" - the content PUT itself must never be attempted (PRD §4.4:
        // parents are never auto-created).
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Get));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
            {
                Content = new StringContent(
                    """{"error":{"code":"itemNotFound","message":"Item not found"}}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });

        Assert.That(
            () => CreateClient(handler).CreateFileAsync("sub/notes.md", "content", conflictBehavior: "replace", driveId: "drive-id"),
            Throws.InstanceOf<DriveParentNotFoundException>());
    }

    [Test]
    public async Task CreateFileAsync_rename_appends_a_numeric_suffix_on_collision()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                // The original path (no "(1)" suffix) is reported as existing; the
                // suffixed candidate is reported as free.
                var isSuffixedCandidate = request.RequestUri!.ToString().Contains("(1)");

                return Task.FromResult(new HttpResponseMessage(isSuffixedCandidate ? HttpStatusCode.NotFound : HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        isSuffixedCandidate
                            ? """{"error":{"code":"itemNotFound","message":"Item not found"}}"""
                            : """{"id":"existing-id"}""",
                        Encoding.UTF8,
                        "application/json"),
                });
            }

            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Put));
                Assert.That(request.RequestUri!.ToString(), Does.Contain("(1)"), "the PUT should target the suffixed, non-colliding path.");
            });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"renamed-id"}""", Encoding.UTF8, "application/json"),
            });
        });

        var result = await CreateClient(handler).CreateFileAsync("notes.md", "content", conflictBehavior: "rename", driveId: "drive-id");

        Assert.That(result?.Id, Is.EqualTo("renamed-id"));
    }

    [Test]
    public async Task GetItemAsync_routes_a_path_looking_string_to_the_colon_path_endpoint()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.RequestUri!.ToString(), Does.Contain("root:/notes.md:"));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"item-id"}""", Encoding.UTF8, "application/json"),
            });
        });

        var resolution = await CreateClient(handler).GetItemAsync("notes.md", driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(resolution.ResolvedAsId, Is.False);
            Assert.That(resolution.Item?.Id, Is.EqualTo("item-id"));
        });
    }

    [Test]
    public async Task GetItemAsync_routes_an_id_looking_string_to_the_items_by_id_endpoint()
    {
        const string itemId = "0176NADPLJLSYNNMETHFBJ7IQMXSGTZ4MA";
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.Multiple(() =>
            {
                Assert.That(request.RequestUri!.ToString(), Does.Contain($"items/{itemId}"));
                Assert.That(request.RequestUri!.ToString(), Does.Not.Contain("root:"));
            });

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent($$"""{"id":"{{itemId}}"}""", Encoding.UTF8, "application/json"),
            });
        });

        var resolution = await CreateClient(handler).GetItemAsync(itemId, driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(resolution.ResolvedAsId, Is.True);
            Assert.That(resolution.Item?.Id, Is.EqualTo(itemId));
        });
    }

    [Test]
    public async Task RenameItemAsync_sends_a_PATCH_with_the_new_name()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            Assert.Multiple(() =>
            {
                Assert.That(request.Method, Is.EqualTo(HttpMethod.Patch));
                Assert.That(request.RequestUri!.ToString(), Does.Contain("items/item-id"));
            });

            var body = await request.Content!.ReadAsStringAsync();
            Assert.That(body, Does.Contain("\"name\":\"new-name.md\""));

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"item-id","name":"new-name.md"}""", Encoding.UTF8, "application/json"),
            };
        });

        var result = await CreateClient(handler).RenameItemAsync("item-id", "new-name.md", driveId: "drive-id");

        Assert.That(result?.Name, Is.EqualTo("new-name.md"));
    }

    [Test]
    public async Task MoveItemAsync_sends_a_PATCH_with_the_new_parent_id()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            if (request.Method == HttpMethod.Get)
            {
                // Destination-parent lookup, used for the same-drive defense-in-depth check.
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        """{"id":"new-parent-id","parentReference":{"driveId":"drive-id"}}""",
                        Encoding.UTF8,
                        "application/json"),
                };
            }

            Assert.That(request.Method, Is.EqualTo(HttpMethod.Patch));

            var body = await request.Content!.ReadAsStringAsync();
            Assert.That(body, Does.Contain("\"id\":\"new-parent-id\""));

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""{"id":"item-id"}""", Encoding.UTF8, "application/json"),
            };
        });

        var result = await CreateClient(handler).MoveItemAsync("item-id", "new-parent-id", driveId: "drive-id");

        Assert.That(result?.Id, Is.EqualTo("item-id"));
    }

    [Test]
    public void MoveItemAsync_throws_DriveCrossDriveMoveException_when_the_destination_resolves_to_a_different_drive()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            // Every request in this scenario is the destination-parent lookup, which reports
            // an owning drive different from the one the call was made under - the PATCH
            // itself must never be attempted.
            Assert.That(request.Method, Is.EqualTo(HttpMethod.Get));

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"id":"new-parent-id","parentReference":{"driveId":"a-different-drive-id"}}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });

        Assert.That(
            () => CreateClient(handler).MoveItemAsync("item-id", "new-parent-id", driveId: "drive-id"),
            Throws.InstanceOf<DriveCrossDriveMoveException>());
    }

    [Test]
    public void DeleteItemByIdAsync_swallows_a_404_as_success()
    {
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound)
        {
            Content = new StringContent(
                """{"error":{"code":"itemNotFound","message":"Item not found"}}""",
                Encoding.UTF8,
                "application/json"),
        }));

        // PRD §7 idempotency: deleting an already-deleted item must not surface as an error.
        Assert.That(
            async () => await CreateClient(handler).DeleteItemByIdAsync("item-id", driveId: "drive-id"),
            Throws.Nothing);
    }

    [Test]
    public void ReplaceContentByIdAsync_throws_DriveItemConcurrencyException_on_a_stubbed_412()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            Assert.That(request.Headers.GetValues("If-Match").Single(), Is.EqualTo("stale-etag"));

            return Task.FromResult(new HttpResponseMessage((HttpStatusCode)412)
            {
                Content = new StringContent(
                    """{"error":{"code":"resourceModified","message":"eTag does not match current value"}}""",
                    Encoding.UTF8,
                    "application/json"),
            });
        });

        Assert.That(
            () => CreateClient(handler).ReplaceContentByIdAsync("item-id", "new content", "stale-etag", driveId: "drive-id"),
            Throws.InstanceOf<DriveItemConcurrencyException>());
    }
}
