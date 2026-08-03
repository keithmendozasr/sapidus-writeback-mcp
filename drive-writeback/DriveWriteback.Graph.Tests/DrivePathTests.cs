namespace DriveWriteback.Graph.Tests;

[TestFixture]
[Category("Unit")]
public class DrivePathTests
{
    [TestCase("", "")]
    [TestCase("/", "")]
    [TestCase("root", "")]
    [TestCase("/a/b/", "a/b")]
    [TestCase("a/b", "a/b")]
    public void Normalize_collapses_root_aliases_and_trims_slashes(string path, string expected)
    {
        Assert.That(DrivePath.Normalize(path), Is.EqualTo(expected));
    }

    [Test]
    public void Validate_throws_on_a_disallowed_character_in_a_segment()
    {
        Assert.That(() => DrivePath.Validate("notes/bad:name.md"), Throws.ArgumentException);
    }

    [Test]
    public void Validate_throws_on_a_dot_dot_segment()
    {
        Assert.That(() => DrivePath.Validate("a/../b"), Throws.ArgumentException);
    }

    [Test]
    public void Validate_does_not_throw_on_a_clean_path()
    {
        Assert.That(() => DrivePath.Validate("Shared Documents/notes/2026-07.md"), Throws.Nothing);
    }
}
