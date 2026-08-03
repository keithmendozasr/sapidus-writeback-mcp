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
}
