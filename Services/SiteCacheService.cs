using System.Text.Json;
using System.Text.Json.Serialization;

namespace GraphCli.Services;

public static class SiteCacheService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".graph-cli");
    private static readonly string CachePath = Path.Combine(ConfigDir, "site-cache.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static SiteCache Load()
    {
        if (!File.Exists(CachePath))
            return new SiteCache();

        var json = File.ReadAllText(CachePath);
        return JsonSerializer.Deserialize<SiteCache>(json, JsonOptions) ?? new SiteCache();
    }

    public static void Save(SiteCache cache)
    {
        Directory.CreateDirectory(ConfigDir);
        File.WriteAllText(CachePath, JsonSerializer.Serialize(cache, JsonOptions));
    }

    /// <summary>
    /// Add or update a site in the cache.
    /// </summary>
    public static void Upsert(string id, string name, string? displayName, string? webUrl)
    {
        var cache = Load();
        UpsertInternal(cache, id, name, displayName, webUrl);
        Save(cache);
    }

    /// <summary>
    /// Add or update multiple sites at once.
    /// </summary>
    public static void UpsertMany(IEnumerable<(string Id, string Name, string? DisplayName, string? WebUrl)> sites)
    {
        var cache = Load();
        foreach (var (id, name, displayName, webUrl) in sites)
            UpsertInternal(cache, id, name, displayName, webUrl);
        Save(cache);
    }

    /// <summary>
    /// Search the cache by name or displayName (case-insensitive contains).
    /// </summary>
    public static List<CachedSite> Search(string query, int top = 20)
    {
        var cache = Load();
        return cache.Sites
            .Where(s => s.Name.Contains(query, StringComparison.OrdinalIgnoreCase)
                     || (s.DisplayName?.Contains(query, StringComparison.OrdinalIgnoreCase) ?? false))
            .OrderByDescending(s => s.LastUsed)
            .Take(top)
            .ToList();
    }

    /// <summary>
    /// Resolve a bare site name (e.g. "TSASite") to a full site identifier
    /// (e.g. "contoso.sharepoint.com:/sites/TSASite") using the cache.
    /// Returns null if no match found.
    /// </summary>
    public static string? Resolve(string site)
    {
        var cache = Load();

        // Exact match on name first
        var match = cache.Sites.FirstOrDefault(s =>
            s.Name.Equals(site, StringComparison.OrdinalIgnoreCase));

        // Then try displayName
        match ??= cache.Sites.FirstOrDefault(s =>
            s.DisplayName != null && s.DisplayName.Equals(site, StringComparison.OrdinalIgnoreCase));

        if (match == null)
            return null;

        // Extract host and path from webUrl: https://host/sites/name -> host:/sites/name
        if (match.WebUrl != null && Uri.TryCreate(match.WebUrl, UriKind.Absolute, out var uri))
        {
            match.LastUsed = DateTimeOffset.UtcNow;
            Save(cache);
            return $"{uri.Host}:{uri.AbsolutePath}";
        }

        return null;
    }

    private static void UpsertInternal(SiteCache cache, string id, string name, string? displayName, string? webUrl)
    {
        var existing = cache.Sites.FirstOrDefault(s =>
            s.Id.Equals(id, StringComparison.OrdinalIgnoreCase));

        if (existing != null)
        {
            if (!string.IsNullOrEmpty(name)) existing.Name = name;
            if (!string.IsNullOrEmpty(displayName)) existing.DisplayName = displayName;
            if (!string.IsNullOrEmpty(webUrl)) existing.WebUrl = webUrl;
            existing.LastUsed = DateTimeOffset.UtcNow;
        }
        else
        {
            cache.Sites.Add(new CachedSite
            {
                Id = id,
                Name = name,
                DisplayName = displayName,
                WebUrl = webUrl,
                LastUsed = DateTimeOffset.UtcNow
            });
        }
    }
}

public class SiteCache
{
    public List<CachedSite> Sites { get; set; } = [];
}

public class CachedSite
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string? DisplayName { get; set; }
    public string? WebUrl { get; set; }
    public DateTimeOffset LastUsed { get; set; }
}
