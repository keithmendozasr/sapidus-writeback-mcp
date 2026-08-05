using System.Text;
using Sapidus.Writeback.Shared.Confirmation;
using Sapidus.Writeback.Shared.Tests.TestSupport;

namespace Sapidus.Writeback.Shared.Tests.Confirmation;

[TestFixture]
[Category("Unit")]
public class ConfirmationTokenServiceTests
{
    private static ConfirmationTokenService CreateService(DateTimeOffset now) =>
        new(Encoding.UTF8.GetBytes("test-signing-key"), new FakeTimeProvider(now));

    [Test]
    public void Validate_accepts_a_token_just_issued_for_the_same_resource_id()
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        var token = service.Issue("fake-resource-id");

        Assert.That(service.Validate("fake-resource-id", token), Is.True);
    }

    [Test]
    public void Validate_rejects_a_token_issued_for_a_different_resource_id()
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        var token = service.Issue("fake-resource-id");

        Assert.That(service.Validate("a-different-resource-id", token), Is.False);
    }

    [Test]
    public void Validate_rejects_a_token_signed_with_a_different_key()
    {
        var issuingService = new ConfirmationTokenService(
            Encoding.UTF8.GetBytes("key-one"),
            new FakeTimeProvider(DateTimeOffset.UtcNow));

        var validatingService = new ConfirmationTokenService(
            Encoding.UTF8.GetBytes("key-two"),
            new FakeTimeProvider(DateTimeOffset.UtcNow));

        var token = issuingService.Issue("fake-resource-id");

        Assert.That(validatingService.Validate("fake-resource-id", token), Is.False);
    }
}
