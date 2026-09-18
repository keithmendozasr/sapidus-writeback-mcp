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

    [Test]
    public async Task RunAsync_rejects_more_than_the_max_batch_size_without_calling_the_service()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph must not be called."));
        var tool = CreateTool(handler);
        var tooManyIds = Enumerable.Range(0, EventDeletionService.MaxBatchSize + 1).Select(i => $"event-{i}").ToArray();

        var result = await tool.RunAsync(null!, tooManyIds, null);

        Assert.That(result, Does.Contain(EventDeletionService.MaxBatchSize.ToString()));
    }

    [Test]
    public async Task RunAsync_reports_when_the_confirmation_token_is_invalid()
    {
        var handler = new StubHttpMessageHandler(_ => throw new InvalidOperationException("Graph must not be called."));
        var tool = CreateTool(handler);

        var result = await tool.RunAsync(null!, ["event-1"], "not-a-valid-token");

        Assert.That(result, Does.Contain("invalid or expired"));
    }
}
