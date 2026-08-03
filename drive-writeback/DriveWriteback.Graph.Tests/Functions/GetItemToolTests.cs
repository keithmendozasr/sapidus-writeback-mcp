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
                """{"id":"folder-id","name":"Reports","folder":{},"eTag":"\"e1\"","cTag":"\"c1\"","size":0,"webUrl":"https://example/Reports"}""",
                Encoding.UTF8,
                "application/json"),
        }));
        var tool = new GetItemTool(CreateClient(handler));

        var result = await tool.RunAsync(null!, "Reports", driveId: "drive-id");

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
                $$"""{"id":"{{itemId}}","name":"notes.md","file":{},"eTag":"\"e1\"","cTag":"\"c1\"","size":42,"webUrl":"https://example/notes.md"}""",
                Encoding.UTF8,
                "application/json"),
        }));
        var tool = new GetItemTool(CreateClient(handler));

        var result = await tool.RunAsync(null!, itemId, driveId: "drive-id");

        Assert.Multiple(() =>
        {
            Assert.That(result, Does.Contain("file"));
            Assert.That(result, Does.Contain($"resolved \"{itemId}\" as ID"));
            Assert.That(result, Does.Contain("42 bytes"));
        });
    }
}
