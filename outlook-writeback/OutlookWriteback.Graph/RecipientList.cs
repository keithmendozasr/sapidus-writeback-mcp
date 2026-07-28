namespace OutlookWriteback.Graph;

/// <summary>
/// Shared normalization/validation for the array-typed recipient fields this server has -
/// create_draft/update_draft's to/cc/bcc and update_event's attendees. Deliberately does not
/// validate email-address format (see prd-multi-recipient.md §3); it only enforces the
/// structural invariant that a populated array can never silently collapse into the "clear"
/// state through a blank entry.
/// </summary>
public static class RecipientList
{
    public static IReadOnlyList<string>? Normalize(IReadOnlyList<string>? raw)
    {
        if (raw is null || raw.Count == 0)
            return raw;

        var normalized = new List<string>(raw.Count);

        foreach (var entry in raw)
        {
            var trimmed = entry.Trim();

            if (trimmed.Length == 0)
                throw new ArgumentException("Recipient list entries must not be blank.");

            normalized.Add(trimmed);
        }

        return normalized;
    }

    /// <summary>
    /// Returns any address present in more than one of the three lists (case-insensitive, since
    /// email addresses are conventionally treated that way), so callers can be asked to
    /// de-duplicate rather than silently keeping the address in only one list. Duplicates within
    /// a single list are not flagged - only cross-list overlap is in scope here.
    /// </summary>
    public static IReadOnlyList<string> FindDuplicates(
        IReadOnlyList<string> to,
        IReadOnlyList<string> cc,
        IReadOnlyList<string> bcc)
    {
        var toSet = new HashSet<string>(to, StringComparer.OrdinalIgnoreCase);
        var ccSet = new HashSet<string>(cc, StringComparer.OrdinalIgnoreCase);
        var bccSet = new HashSet<string>(bcc, StringComparer.OrdinalIgnoreCase);

        var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var duplicates = new List<string>();

        foreach (var address in toSet.Concat(ccSet).Concat(bccSet))
        {
            if (!reported.Add(address))
                continue;

            var listCount = (toSet.Contains(address) ? 1 : 0)
                + (ccSet.Contains(address) ? 1 : 0)
                + (bccSet.Contains(address) ? 1 : 0);

            if (listCount > 1)
                duplicates.Add(address);
        }

        return duplicates;
    }
}
