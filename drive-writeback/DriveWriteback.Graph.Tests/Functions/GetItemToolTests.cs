using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using DriveWriteback.Functions;
using DriveWriteback.Graph;
using DriveWriteback.Graph.Tests.TestSupport;

namespace DriveWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class GetItemToolTests
{
    private const string NotModifiedSince = "2026-09-01T12:00:00Z";

    private static DriveGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new DriveGraphClient(graphClient);
    }

    [Test]
    public async Task RunAsync_reports_folder_kind_and_that_the_input_was_resolved_as_a_path()
    {
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(
                $$"""{"id":"folder-id","name":"Reports","folder":{},"eTag":"\"e1\"","cTag":"\"c1\"","size":0,"webUrl":"https://example/Reports","lastModifiedDateTime":"{{NotModifiedSince}}"}""",
                Encoding.UTF8,
                "application/json"),
        }));
        var tool = new GetItemTool(CreateClient(handler), new RecordingLogger<GetItemTool>());

        var result = await tool.RunAsync(null!, "Reports", NotModifiedSince, driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("folder"));
            Assert.That(result, Does.Contain("resolved \"Reports\" as path"));
            Assert.That(result, Does.Contain("folder-id"));
        });
    }

    [Test]
    public async Task RunAsync_reports_file_kind_and_that_the_input_was_resolved_as_an_id()
    {
        const string itemId = "0176NADPLJLSYNNMETHFBJ7IQMXSGTZ4MA";
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(
                $$"""{"id":"{{itemId}}","name":"notes.md","file":{},"eTag":"\"e1\"","cTag":"\"c1\"","size":42,"webUrl":"https://example/notes.md","lastModifiedDateTime":"{{NotModifiedSince}}"}""",
                Encoding.UTF8,
                "application/json"),
        }));
        var tool = new GetItemTool(CreateClient(handler), new RecordingLogger<GetItemTool>());

        var result = await tool.RunAsync(null!, itemId, NotModifiedSince, driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("file"));
            Assert.That(result, Does.Contain($"resolved \"{itemId}\" as ID"));
            Assert.That(result, Does.Contain("42 bytes"));
        });
    }

    [Test]
    public async Task RunAsync_surfaces_the_if_match_value_with_its_quote_characters_intact()
    {
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(
                $$"""{"id":"item-id","name":"notes.md","file":{},"eTag":"\"{GUID},1\"","cTag":"\"c:{GUID},1\"","size":1,"webUrl":"https://example/notes.md","lastModifiedDateTime":"{{NotModifiedSince}}"}""",
                Encoding.UTF8,
                "application/json"),
        }));
        var tool = new GetItemTool(CreateClient(handler), new RecordingLogger<GetItemTool>());

        var result = await tool.RunAsync(null!, "notes.md", NotModifiedSince, driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("\"{GUID},1\""), "the eTag's surrounding quotes must survive into the response verbatim - a model copying an unquoted substring produces an if_match value Graph will always reject with a 412, no matter how fresh the read.");
            Assert.That(result, Does.Contain("\"c:{GUID},1\""));
            Assert.That(result, Does.Contain("copy this exact string verbatim"));
        });
    }
}
