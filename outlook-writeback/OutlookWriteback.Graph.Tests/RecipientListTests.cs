using OutlookWriteback.Graph;

namespace OutlookWriteback.Graph.Tests;

[TestFixture]
[Category("Unit")]
public class RecipientListTests
{
    [Test]
    public void Normalize_passes_null_through_unchanged()
    {
        Assert.That(RecipientList.Normalize(null), Is.Null);
    }

    [Test]
    public void Normalize_passes_empty_array_through_unchanged()
    {
        var result = RecipientList.Normalize([]);

        Assert.That(result, Is.Not.Null.And.Empty);
    }

    [Test]
    public void Normalize_trims_whitespace_around_entries()
    {
        var result = RecipientList.Normalize(["  alice@example.com ", "bob@example.com\t"]);

        Assert.That(result, Is.EqualTo(new[] { "alice@example.com", "bob@example.com" }));
    }

    [Test]
    public void Normalize_throws_when_a_populated_array_contains_a_blank_entry()
    {
        Assert.That(
            () => RecipientList.Normalize(["alice@example.com", "   "]),
            Throws.ArgumentException);
    }
}
