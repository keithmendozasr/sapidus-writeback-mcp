using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Functions;
using OutlookWriteback.Graph;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class CreateDraftToolTests
{
    private static OutlookGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new OutlookGraphClient(graphClient);
    }

    [Test]
    public void RunAsync_throws_when_a_recipient_entry_is_blank()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph should not be called."));
        var tool = new CreateDraftTool(CreateClient(handler));

        Assert.That(
            () => tool.RunAsync(null!, ["alice@example.com", "   "], "Subject", "Body", null, null, null),
            Throws.ArgumentException);
    }

    [Test]
    public void RunAsync_throws_when_the_same_address_appears_in_more_than_one_list()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph should not be called."));
        var tool = new CreateDraftTool(CreateClient(handler));

        Assert.That(
            () => tool.RunAsync(null!, ["alice@example.com"], "Subject", "Body", null, ["alice@example.com"], null),
            Throws.ArgumentException);
    }

    [Test]
    public async Task RunAsync_trims_whitespace_before_sending_to_graph()
    {
        var handler = new StubHttpMessageHandler(async request =>
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync();

            Assert.That(body, Does.Contain("\"address\":\"alice@example.com\""));

            return new HttpResponseMessage(System.Net.HttpStatusCode.Created)
            {
                Content = new StringContent("""{"id":"AAMk-fake-draft-id"}""", Encoding.UTF8, "application/json"),
            };
        });
        var tool = new CreateDraftTool(CreateClient(handler));

        await tool.RunAsync(null!, ["  alice@example.com  "], "Subject", "Body", null, null, null);
    }

    [Test]
    public async Task RunAsync_creates_a_draft_with_no_recipients_when_to_cc_bcc_are_all_omitted()
    {
        var handler = new StubHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(System.Net.HttpStatusCode.Created)
        {
            Content = new StringContent("""{"id":"AAMk-fake-draft-id"}""", Encoding.UTF8, "application/json"),
        }));
        var tool = new CreateDraftTool(CreateClient(handler));

        var result = await tool.RunAsync(null!, null, "Subject", "Body", null, null, null);

        Assert.That(result, Does.Contain("AAMk-fake-draft-id"));
    }
}
