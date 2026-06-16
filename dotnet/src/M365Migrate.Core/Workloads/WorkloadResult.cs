namespace M365Migrate.Core.Workloads;

/// <summary>
/// A per-item outcome from a migrate/sync run. <see cref="Name"/> identifies the
/// item (e.g. a UPN or mailNickname); <see cref="Detail"/> holds workload-specific
/// sub-statuses (members, owners, grants, ...).
/// </summary>
public sealed class WorkloadResult
{
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
    public string? Reason { get; set; }
    public Dictionary<string, string> Detail { get; set; } = new();

    public WorkloadResult() { }

    public WorkloadResult(string name, string status, string? reason = null)
    {
        Name = name;
        Status = status;
        Reason = reason;
    }
}
