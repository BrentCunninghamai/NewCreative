using System.Text.Json;

namespace M365Migrate.Core.Models;

/// <summary>
/// A mailbox-settings migration planned for the target tenant. <see cref="Action"/>
/// is one of <c>settings</c> (apply) or <c>skip</c>.
/// </summary>
public sealed class PlannedMailbox
{
    public string SourceUpn { get; set; } = "";
    public string TargetUpn { get; set; } = "";
    public string Action { get; set; } = "settings";
    public string? Reason { get; set; }
    public Dictionary<string, JsonElement> Settings { get; set; } = new();
}
