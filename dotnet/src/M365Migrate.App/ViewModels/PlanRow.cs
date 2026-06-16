namespace M365Migrate.App.ViewModels;

/// <summary>A single row shown in the plan/results grid.</summary>
public sealed class PlanRow
{
    public string Name { get; init; } = "";
    public string Action { get; init; } = "";
    public string Detail { get; init; } = "";
    public string Reason { get; init; } = "";
}
