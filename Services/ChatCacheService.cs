using System.Text.Json;
using System.Text.Json.Serialization;

namespace GraphCli.Services;

public static class ChatCacheService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".graph-cli");
    private static readonly string CachePath = Path.Combine(ConfigDir, "chat-cache.json");

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
    /// Add or update a chat in the cache.
    /// </summary>
    public static void Upsert(string id, string? name, string? chatType)
    {
        var cache = Load();
        var existing = cache.Chats.FirstOrDefault(c =>
            c.Id.Equals(id, StringComparison.OrdinalIgnoreCase));

        if (existing != null)
        {
            if (!string.IsNullOrEmpty(name)) existing.Name = name;
            if (!string.IsNullOrEmpty(chatType)) existing.ChatType = chatType;
            existing.LastUsed = DateTimeOffset.UtcNow;
        }
        else
        {
            cache.Chats.Add(new CachedChat
            {
                Id = id,
                Name = name ?? "",
                ChatType = chatType ?? "unknown",
                LastUsed = DateTimeOffset.UtcNow
            });
        }

        Save(cache);
    }

    /// <summary>
    /// Add or update multiple chats at once.
    /// </summary>
    public static void UpsertMany(IEnumerable<(string Id, string? Name, string? ChatType)> chats)
    {
        var cache = Load();

        foreach (var (id, name, chatType) in chats)
        {
            var existing = cache.Chats.FirstOrDefault(c =>
                c.Id.Equals(id, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                if (!string.IsNullOrEmpty(name)) existing.Name = name;
                if (!string.IsNullOrEmpty(chatType)) existing.ChatType = chatType;
                existing.LastUsed = DateTimeOffset.UtcNow;
            }
            else
            {
                cache.Chats.Add(new CachedChat
                {
                    Id = id,
                    Name = name ?? "",
                    ChatType = chatType ?? "unknown",
                    LastUsed = DateTimeOffset.UtcNow
                });
            }
        }

        Save(cache);
    }

    /// <summary>
    /// Search the cache by name (case-insensitive contains).
    /// </summary>
    public static List<CachedChat> Search(string query, int top = 20)
    {
        var cache = Load();
        return cache.Chats
            .Where(c => c.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(c => c.LastUsed)
            .Take(top)
            .ToList();
    }
}

public class ChatCache
{
    public List<CachedChat> Chats { get; set; } = [];
}

public class CachedChat
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string ChatType { get; set; } = "unknown";
    public DateTimeOffset LastUsed { get; set; }
}
