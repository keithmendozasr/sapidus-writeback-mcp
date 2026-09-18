using System.Net;
using System.Text;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookWriteback.Functions;
using OutlookWriteback.Graph.Confirmation;
using OutlookWriteback.Graph.Tests.TestSupport;
using Sapidus.Writeback.Shared.Confirmation;

namespace OutlookWriteback.Graph.Tests.Functions;

[TestFixture]
[Category("Unit")]
public class DeleteEventToolTests
{
    private static DeleteEventTool CreateTool(StubHttpMessageHandler handler, ConfirmationTokenService? tokenService = null)
    {
        var httpClient = new HttpClient(handler) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0") };
        var graphClient = new OutlookGraphClient(new GraphServiceClient(httpClient, new AnonymousAuthenticationProvider()));
        var deletionService = new EventDeletionService(graphClient, tokenService ?? new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key")));

        return new DeleteEventTool(deletionService);
    }

    [Test]
    public async Task RunAsync_requires_at_least_one_event_id()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph must not be called."));
        var tool = CreateTool(handler);

        var result = await tool.RunAsync(null!, [], null);

        Assert.That(result, Is.EqualTo("At least one eventId is required."));
    }
}
