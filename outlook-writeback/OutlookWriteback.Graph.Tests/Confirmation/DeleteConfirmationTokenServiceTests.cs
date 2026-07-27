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

    [Test]
    public void Validate_rejects_a_token_issued_for_a_different_event_id()
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        var token = service.Issue("AAkA-fake-event-id");

        Assert.That(service.Validate("AAkA-a-different-event-id", token), Is.False);
    }

    [Test]
    public void Validate_rejects_a_token_signed_with_a_different_key()
    {
        var issuingService = new DeleteConfirmationTokenService(
            Encoding.UTF8.GetBytes("key-one"),
            new FakeTimeProvider(DateTimeOffset.UtcNow));

        var validatingService = new DeleteConfirmationTokenService(
            Encoding.UTF8.GetBytes("key-two"),
            new FakeTimeProvider(DateTimeOffset.UtcNow));

        var token = issuingService.Issue("AAkA-fake-event-id");

        Assert.That(validatingService.Validate("AAkA-fake-event-id", token), Is.False);
    }

    [Test]
    public void Validate_rejects_a_token_whose_payload_has_been_tampered_with()
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        var token = service.Issue("AAkA-fake-event-id");
        var payload = token.Split('.')[0];
        var tampered = token.Replace(payload, payload + "x");

        Assert.That(service.Validate("AAkA-fake-event-id", tampered), Is.False);
    }

    [Test]
    public void Validate_rejects_a_token_after_its_ttl_has_elapsed()
    {
        var timeProvider = new FakeTimeProvider(DateTimeOffset.UtcNow);
        var service = new DeleteConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), timeProvider);

        var token = service.Issue("AAkA-fake-event-id");
        timeProvider.Now = timeProvider.Now.AddMinutes(6);

        Assert.That(service.Validate("AAkA-fake-event-id", token), Is.False);
    }

    [Test]
    [TestCase("")]
    [TestCase("not-a-token")]
    [TestCase("too.many.parts")]
    [TestCase("not-valid-base64!.also-not-valid-base64!")]
    public void Validate_rejects_malformed_tokens_without_throwing(string malformedToken)
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        Assert.That(service.Validate("AAkA-fake-event-id", malformedToken), Is.False);
    }
}
