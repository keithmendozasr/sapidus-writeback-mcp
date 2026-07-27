using System.Text;
using OutlookWriteback.Graph.Confirmation;
using OutlookWriteback.Graph.Tests.TestSupport;

namespace OutlookWriteback.Graph.Tests.Confirmation;

[TestFixture]
[Category("Unit")]
public class DeleteConfirmationTokenServiceTests
{
    private static DeleteConfirmationTokenService CreateService(DateTimeOffset now) =>
        new(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));

    [Test]
    public void Validate_accepts_a_token_just_issued_for_the_same_event_id()
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        var token = service.Issue("AAkA-fake-event-id");

        Assert.That(service.Validate("AAkA-fake-event-id", token), Is.True);
    }
}
