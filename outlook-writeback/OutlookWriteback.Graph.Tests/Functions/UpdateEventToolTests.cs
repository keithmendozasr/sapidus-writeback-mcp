using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Functions;
using OutlookWriteback.Graph;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class UpdateEventToolTests
{
    private static OutlookGraphClient CreateClient(StubHttpMessageHandler handler)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider());

        return new OutlookGraphClient(graphClient);
    }

    [Test]
    public void RunAsync_throws_when_an_attendee_entry_is_blank()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph should not be called."));
        var tool = new UpdateEventTool(CreateClient(handler));

        Assert.That(
            () => tool.RunAsync(null!, "AAkA-fake-event-id", null, null, null, null, null, ["alice@example.com", "   "]),
            Throws.ArgumentException);
    }
}
