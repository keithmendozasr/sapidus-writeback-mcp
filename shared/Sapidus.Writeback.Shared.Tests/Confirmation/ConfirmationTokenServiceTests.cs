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

    [Test]
    public void Validate_rejects_a_token_whose_payload_has_been_tampered_with()
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        var token = service.Issue("fake-resource-id");
        var payload = token.Split('.')[0];
        var tampered = token.Replace(payload, payload + "x");

        Assert.That(service.Validate("fake-resource-id", tampered), Is.False);
    }

    [Test]
    public void Validate_rejects_a_token_after_its_ttl_has_elapsed()
    {
        var timeProvider = new FakeTimeProvider(DateTimeOffset.UtcNow);
        var service = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), timeProvider);

        var token = service.Issue("fake-resource-id");
        timeProvider.Now = timeProvider.Now.AddMinutes(6);

        Assert.That(service.Validate("fake-resource-id", token), Is.False);
    }

    [Test]
    [TestCase("")]
    [TestCase("not-a-token")]
    [TestCase("too.many.parts")]
    [TestCase("not-valid-base64!.also-not-valid-base64!")]
    public void Validate_rejects_malformed_tokens_without_throwing(string malformedToken)
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        Assert.That(service.Validate("fake-resource-id", malformedToken), Is.False);
    }

    [Test]
    public void ValidateSubset_authorizes_every_previewed_id_when_the_full_set_is_reconfirmed()
    {
        var service = CreateService(DateTimeOffset.UtcNow);
        var ids = new[] { "event-1", "event-2", "event-3" };

        var token = service.IssueBatch(ids);
        var result = service.ValidateSubset(ids, token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.True);
            Assert.That(result.AuthorizedIds, Is.EquivalentTo(ids));
        });
    }

    [Test]
    public void ValidateSubset_authorizes_only_the_resent_subset_when_some_ids_are_dropped()
    {
        var service = CreateService(DateTimeOffset.UtcNow);
        var token = service.IssueBatch(["event-1", "event-2", "event-3"]);

        var result = service.ValidateSubset(["event-1", "event-3"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.True);
            Assert.That(result.AuthorizedIds, Is.EquivalentTo(new[] { "event-1", "event-3" }));
        });
    }

    [Test]
    public void ValidateSubset_does_not_authorize_an_id_outside_the_originally_issued_set()
    {
        var service = CreateService(DateTimeOffset.UtcNow);
        var token = service.IssueBatch(["event-1", "event-2"]);

        var result = service.ValidateSubset(["event-1", "not-previewed"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.True);
            Assert.That(result.AuthorizedIds, Is.EquivalentTo(new[] { "event-1" }));
        });
    }

    [Test]
    public void ValidateSubset_rejects_the_whole_token_when_signed_with_a_different_key()
    {
        var issuingService = new ConfirmationTokenService(
            Encoding.UTF8.GetBytes("key-one"),
            new FakeTimeProvider(DateTimeOffset.UtcNow));

        var validatingService = new ConfirmationTokenService(
            Encoding.UTF8.GetBytes("key-two"),
            new FakeTimeProvider(DateTimeOffset.UtcNow));

        var token = issuingService.IssueBatch(["event-1", "event-2"]);
        var result = validatingService.ValidateSubset(["event-1"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.False);
            Assert.That(result.AuthorizedIds, Is.Empty);
        });
    }

    [Test]
    public void ValidateSubset_rejects_the_whole_token_when_tampered_with()
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        var token = service.IssueBatch(["event-1", "event-2"]);
        var payload = token.Split('.')[0];
        var tampered = token.Replace(payload, payload + "x");

        Assert.That(service.ValidateSubset(["event-1"], tampered).TokenValid, Is.False);
    }

    [Test]
    public void ValidateSubset_rejects_the_whole_token_after_its_ttl_has_elapsed()
    {
        var timeProvider = new FakeTimeProvider(DateTimeOffset.UtcNow);
        var service = new ConfirmationTokenService(Encoding.UTF8.GetBytes("test-signing-key"), timeProvider);

        var token = service.IssueBatch(["event-1", "event-2"]);
        timeProvider.Now = timeProvider.Now.AddMinutes(6);

        Assert.That(service.ValidateSubset(["event-1"], token).TokenValid, Is.False);
    }

    [Test]
    [TestCase("")]
    [TestCase("not-a-token")]
    [TestCase("too.many.parts")]
    [TestCase("not-valid-base64!.also-not-valid-base64!")]
    public void ValidateSubset_rejects_malformed_tokens_without_throwing(string malformedToken)
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        Assert.That(service.ValidateSubset(["event-1"], malformedToken).TokenValid, Is.False);
    }

    [Test]
    public void IssueBatch_and_ValidateSubset_round_trip_regardless_of_id_order()
    {
        var service = CreateService(DateTimeOffset.UtcNow);

        var token = service.IssueBatch(["event-3", "event-1", "event-2"]);
        var result = service.ValidateSubset(["event-2", "event-1"], token);

        Assert.Multiple(() =>
        {
            Assert.That(result.TokenValid, Is.True);
            Assert.That(result.AuthorizedIds, Is.EquivalentTo(new[] { "event-1", "event-2" }));
        });
    }
}
