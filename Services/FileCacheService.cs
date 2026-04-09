using System.Text.Json;
using System.Text.Json.Serialization;

namespace GraphCli.Services;

public static class FileCacheService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".graph-cli");
    private static readonly string CachePath = Path.Combine(ConfigDir, "file-cache.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static FileCache Load()
    {
        if (!File.Exists(CachePath))
            return new FileCache();

        var json = File.ReadAllText(CachePath);
        return JsonSerializer.Deserialize<FileCache>(json, JsonOptions) ?? new FileCache();
    }

    public static void Save(FileCache cache)
    {
        Directory.CreateDirectory(ConfigDir);
        File.WriteAllText(CachePath, JsonSerializer.Serialize(cache, JsonOptions));
    }

    /// <summary>
    /// Add or update a file in the cache.
    /// </summary>
    public static void Upsert(string id, string name, string type, long? size, string? mimeType, string? webUrl)
    {
        var cache = Load();
        UpsertInternal(cache, id, name, type, size, mimeType, webUrl);
        Save(cache);
    }

    /// <summary>
    /// Add or update multiple files at once.
    /// </summary>
    public static void UpsertMany(IEnumerable<(string Id, string Name, string Type, long? Size, string? MimeType, string? WebUrl)> files)
    {
        var cache = Load();
        foreach (var (id, name, type, size, mimeType, webUrl) in files)
            UpsertInternal(cache, id, name, type, size, mimeType, webUrl);
        Save(cache);
    }

    /// <summary>
    /// Search the cache by name (case-insensitive contains).
    /// </summary>
    public static List<CachedFile> Search(string query, int top = 25)
    {
        var cache = Load();
        return cache.Files
            .Where(f => f.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(f => f.LastUsed)
            .Take(top)
            .ToList();
    }

    private static void UpsertInternal(FileCache cache, string id, string name, string type, long? size, string? mimeType, string? webUrl)
    {
        var existing = cache.Files.FirstOrDefault(f =>
            f.Id.Equals(id, StringComparison.OrdinalIgnoreCase));

        if (existing != null)
        {
            if (!string.IsNullOrEmpty(name)) existing.Name = name;
            if (!string.IsNullOrEmpty(type)) existing.Type = type;
            if (size.HasValue) existing.Size = size;
            if (!string.IsNullOrEmpty(mimeType)) existing.MimeType = mimeType;
            if (!string.IsNullOrEmpty(webUrl)) existing.WebUrl = webUrl;
            existing.LastUsed = DateTimeOffset.UtcNow;
        }
        else
        {
            cache.Files.Add(new CachedFile
            {
                Id = id,
                Name = name,
                Type = type,
                Size = size,
                MimeType = mimeType,
                WebUrl = webUrl,
                LastUsed = DateTimeOffset.UtcNow
            });
        }
    }
}

public class FileCache
{
    public List<CachedFile> Files { get; set; } = [];
}

public class CachedFile
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Type { get; set; } = "file";
    public long? Size { get; set; }
    public string? MimeType { get; set; }
    public string? WebUrl { get; set; }
    public DateTimeOffset LastUsed { get; set; }
}
