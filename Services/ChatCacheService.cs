using System.Text.Json;
using System.Text.Json.Serialization;

namespace GraphCli.Services;

public static class ChatCacheService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".graph-cli");
    private static readonly string CachePath = Path.Combine(ConfigDir, "chat-cache.json");
    private static readonly string WatermarkPath = Path.Combine(ConfigDir, "chat-since-watermark.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static ChatCache Load()
    {
        if (!File.Exists(CachePath))
            return new ChatCache();

        var json = File.ReadAllText(CachePath);
        return JsonSerializer.Deserialize<ChatCache>(json, JsonOptions) ?? new ChatCache();
    }

    public static void Save(ChatCache cache)
    {
        Directory.CreateDirectory(ConfigDir);
        File.WriteAllText(CachePath, JsonSerializer.Serialize(cache, JsonOptions));
    }

    /// <summary>
    /// Add or update a chat in the cache. Topic-only overload kept for callers
    /// that don't have member or lastUpdated info (e.g. chat get / chat list).
    /// </summary>
    public static void Upsert(string id, string? name, string? chatType)
        => UpsertMany(new[] { new ChatCacheEntry(id, name, chatType, null, null) });

    /// <summary>
    /// Topic-only batch upsert. Kept for back-compat with chat list / find-with.
    /// </summary>
    public static void UpsertMany(IEnumerable<(string Id, string? Name, string? ChatType)> chats)
        => UpsertMany(chats.Select(c => new ChatCacheEntry(c.Id, c.Name, c.ChatType, null, null)));

    /// <summary>
    /// Full upsert with members and lastUpdated. Existing fields are preserved
    /// when the new entry's value is null (e.g. a topic-only call won't wipe
    /// previously cached members).
    /// </summary>
    public static void UpsertMany(IEnumerable<ChatCacheEntry> entries)
    {
        var cache = Load();

        foreach (var e in entries)
        {
            var existing = cache.Chats.FirstOrDefault(c =>
                c.Id.Equals(e.Id, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                if (!string.IsNullOrEmpty(e.Name)) existing.Name = e.Name;
                if (!string.IsNullOrEmpty(e.ChatType)) existing.ChatType = e.ChatType;
                if (e.Members != null) existing.Members = e.Members;
                if (e.LastUpdatedDateTime.HasValue) existing.LastUpdatedDateTime = e.LastUpdatedDateTime;
                existing.LastUsed = DateTimeOffset.UtcNow;
            }
            else
            {
                cache.Chats.Add(new CachedChat
                {
                    Id = e.Id,
                    Name = e.Name ?? "",
                    ChatType = e.ChatType ?? "unknown",
                    Members = e.Members,
                    LastUpdatedDateTime = e.LastUpdatedDateTime,
                    LastUsed = DateTimeOffset.UtcNow
                });
            }
        }

        Save(cache);
    }

    /// <summary>
    /// Search the cache by name OR member display name OR member email (case-insensitive contains).
    /// </summary>
    public static List<CachedChat> Search(string query, int top = 20)
    {
        var cache = Load();
        return cache.Chats
            .Where(c => Matches(c, query))
            .OrderByDescending(c => c.LastUpdatedDateTime ?? c.LastUsed)
            .Take(top)
            .ToList();
    }

    /// <summary>
    /// Read the persisted "since" watermark — the createdDateTime of the newest chat
    /// message returned by the last `chat since` run. Returns null if none stored yet.
    /// Used so automation can pull only messages that have arrived since the previous poll.
    /// </summary>
    public static DateTimeOffset? GetWatermark()
    {
        if (!File.Exists(WatermarkPath))
            return null;
        try
        {
            var json = File.ReadAllText(WatermarkPath);
            return JsonSerializer.Deserialize<ChatSinceWatermark>(json, JsonOptions)?.Watermark;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Persist the "since" watermark (the newest message timestamp seen this run) so the
    /// next `chat since --continue` picks up exactly where this run left off.
    /// </summary>
    public static void SaveWatermark(DateTimeOffset watermark)
    {
        Directory.CreateDirectory(ConfigDir);
        var payload = new ChatSinceWatermark { Watermark = watermark, UpdatedAt = DateTimeOffset.UtcNow };
        File.WriteAllText(WatermarkPath, JsonSerializer.Serialize(payload, JsonOptions));
    }

    private static bool Matches(CachedChat c, string query)
    {
        if (!string.IsNullOrEmpty(c.Name) && c.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
            return true;

        if (c.Members != null)
        {
            foreach (var m in c.Members)
            {
                if (!string.IsNullOrEmpty(m.DisplayName) && m.DisplayName.Contains(query, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (!string.IsNullOrEmpty(m.Email) && m.Email.Contains(query, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }

        return false;
    }
}

public record ChatCacheEntry(
    string Id,
    string? Name,
    string? ChatType,
    List<CachedMember>? Members,
    DateTimeOffset? LastUpdatedDateTime);

public class ChatCache
{
    public List<CachedChat> Chats { get; set; } = [];
}

public class ChatSinceWatermark
{
    public DateTimeOffset? Watermark { get; set; }
    public DateTimeOffset UpdatedAt { get; set; }
}

public class CachedChat
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string ChatType { get; set; } = "unknown";
    public List<CachedMember>? Members { get; set; }
    public DateTimeOffset? LastUpdatedDateTime { get; set; }
    public DateTimeOffset LastUsed { get; set; }
}

public class CachedMember
{
    public string? DisplayName { get; set; }
    public string? Email { get; set; }
    public string? UserId { get; set; }
}
