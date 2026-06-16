using System.Text.Json;

namespace M365Migrate.Core.Graph;

/// <summary>Null-safe readers for the loosely-typed JSON Graph returns.</summary>
public static class JsonExtensions
{
    public static string? GetStringOrNull(this JsonElement element, string property)
    {
        if (element.ValueKind == JsonValueKind.Object
            && element.TryGetProperty(property, out var value)
            && value.ValueKind == JsonValueKind.String)
            return value.GetString();
        return null;
    }

    public static bool GetBoolOrDefault(this JsonElement element, string property, bool fallback = false)
    {
        if (element.ValueKind == JsonValueKind.Object
            && element.TryGetProperty(property, out var value)
            && (value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False))
            return value.GetBoolean();
        return fallback;
    }

    public static IEnumerable<JsonElement> GetArrayOrEmpty(this JsonElement element, string property)
    {
        if (element.ValueKind == JsonValueKind.Object
            && element.TryGetProperty(property, out var value)
            && value.ValueKind == JsonValueKind.Array)
            return value.EnumerateArray();
        return Array.Empty<JsonElement>();
    }
}
