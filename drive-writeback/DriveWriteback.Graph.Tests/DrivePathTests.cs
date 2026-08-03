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

    [Test]
    public void LooksLikeItemId_is_true_for_the_observed_live_OneDrive_id_shape()
    {
        // 34 chars, alphanumeric, no '/' - the live OneDrive item id observed in
        // Phase 0 (docs/active/PRD-drive-write.md §11).
        Assert.That(DrivePath.LooksLikeItemId("0176NADPLJLSYNNMETHFBJ7IQMXSGTZ4MA"), Is.True);
    }

    [Test]
    public void LooksLikeItemId_is_false_for_a_root_level_filename()
    {
        Assert.That(DrivePath.LooksLikeItemId("notes.md"), Is.False);
    }
}
