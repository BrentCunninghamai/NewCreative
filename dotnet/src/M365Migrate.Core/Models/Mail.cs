using System.Text.Json;
using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Models;

/// <summary>A mail folder as read from a mailbox, with its path relative to the root.</summary>
public sealed class MailFolderInfo
{
    public string Id { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string? ParentFolderId { get; set; }
    public int TotalItemCount { get; set; }
    public int ChildFolderCount { get; set; }

    /// <summary>Folder path relative to the mailbox root, e.g. "Inbox" or "Inbox/2025".</summary>
    public string Path { get; set; } = "";

    public static MailFolderInfo FromGraph(JsonElement data, string parentPath)
    {
        var name = data.GetStringOrNull("displayName") ?? "";
        return new MailFolderInfo
        {
            Id = data.GetStringOrNull("id") ?? "",
            DisplayName = name,
            ParentFolderId = data.GetStringOrNull("parentFolderId"),
            TotalItemCount = data.TryGetProperty("totalItemCount", out var t) && t.ValueKind == JsonValueKind.Number ? t.GetInt32() : 0,
            ChildFolderCount = data.TryGetProperty("childFolderCount", out var c) && c.ValueKind == JsonValueKind.Number ? c.GetInt32() : 0,
            Path = string.IsNullOrEmpty(parentPath) ? name : $"{parentPath}/{name}",
        };
    }
}

/// <summary>A mail folder's content copy planned for the target mailbox.</summary>
public sealed class PlannedMailFolder
{
    public string SourceId { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string Path { get; set; } = "";
    public int ItemCount { get; set; }
    public string Action { get; set; } = "copy";

    /// <summary>The parent folder path ("" for a top-level folder).</summary>
    public string ParentPath
    {
        get
        {
            var slash = Path.LastIndexOf('/');
            return slash < 0 ? "" : Path[..slash];
        }
    }
}
