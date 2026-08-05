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
}
