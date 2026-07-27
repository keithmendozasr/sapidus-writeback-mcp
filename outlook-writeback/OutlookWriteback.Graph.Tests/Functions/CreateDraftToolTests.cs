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
}
