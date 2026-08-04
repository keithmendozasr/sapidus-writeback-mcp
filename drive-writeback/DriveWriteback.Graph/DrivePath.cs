namespace DriveWriteback.Graph;

/// <summary>
/// Pure normalization/validation for the drive-relative addressing model (PRD §5) - no
/// Graph dependency, mirrors RecipientList.cs's role in outlook-writeback.
/// </summary>
public static class DrivePath
{
    private static readonly char[] DisallowedChars = ['"', '*', ':', '<', '>', '?', '\\', '|'];

    public static string Normalize(string path) =>
        path is "" or "/" or "root" ? "" : path.Trim('/');

    /// <summary>
    /// Throws on a disallowed character within any segment (PRD §5's OneDrive/SharePoint
    /// name restrictions) or a ".." segment (no traversal). '/' itself is never checked
    /// here - it's the segment separator, not something that can appear within one once
    /// the path has been split.
    /// </summary>
    public static void Validate(string path)
    {
        var normalized = Normalize(path);

        if (normalized.Length == 0)
            return;

        foreach (var segment in normalized.Split('/'))
        {
            if (segment == "..")
                throw new ArgumentException($"Path segment '..' is not allowed in '{path}'.", nameof(path));

            if (segment.IndexOfAny(DisallowedChars) >= 0)
                throw new ArgumentException($"Path segment '{segment}' contains a disallowed character in '{path}'.", nameof(path));
        }
    }

    /// <summary>
    /// The PRD's stated discriminator ("an id never contains '/'") is insufficient alone -
    /// "notes.md" also contains no '/' either. Adds no-dot/all-alphanumeric/length>=20 (the
    /// one observed live OneDrive item id was 34 chars) so a root-level filename doesn't
    /// misclassify as an id. Callers should still surface which interpretation was used
    /// (see DriveGraphClient.GetItemAsync) so a misclassification is visible, not silent.
    /// </summary>
    public static bool LooksLikeItemId(string pathOrId) =>
        pathOrId.Length >= 20
        && !pathOrId.Contains('/')
        && !pathOrId.Contains('.')
        && pathOrId.All(char.IsLetterOrDigit);

    /// <summary>
    /// "a/b/c.md" -> ("a/b", "c.md"); "c.md" -> ("", "c.md").
    /// </summary>
    public static (string ParentPath, string Name) SplitParent(string path)
    {
        var normalized = Normalize(path);
        var lastSlash = normalized.LastIndexOf('/');

        return lastSlash < 0
            ? ("", normalized)
            : (normalized[..lastSlash], normalized[(lastSlash + 1)..]);
    }
}
