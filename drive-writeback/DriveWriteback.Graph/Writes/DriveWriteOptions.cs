namespace DriveWriteback.Graph.Writes;

/// <summary>
/// Server-level config governing DriveWriteService's behavior (PRD §9 items 5 and 8).
/// DryRun defaults to true so writes are validated-but-not-mutating until explicitly
/// trusted against the live tenant; MaxContentBytes is the create_file size cap,
/// enforced before any Graph call, not left to Graph's own 4 MB simple-upload ceiling.
/// </summary>
public sealed record DriveWriteOptions(bool DryRun, long MaxContentBytes)
{
    public static DriveWriteOptions Default { get; } = new(DryRun: true, MaxContentBytes: 1_048_576);
}
