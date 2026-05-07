using System.Text.Json;
using System.Text.Json.Serialization;

namespace GraphCli.Services;

public static class ConsentService
{
    public const string CurrentVersion = "1";

    public const string ConsentText =
        "graph-cli Terms of Use\n" +
        "\n" +
        "By granting consent you acknowledge that:\n" +
        "  1. graph-cli acts on your behalf against Microsoft Graph using your signed-in\n" +
        "     account and the permissions you authorize.\n" +
        "  2. You understand that commands (and MCP tool calls) may read, modify, send,\n" +
        "     or delete your mail, calendar events, chats, tasks, contacts, and files.\n" +
        "  3. All actions performed through this tool, including those initiated by an AI\n" +
        "     assistant using the MCP server, are solely your responsibility.\n" +
        "  4. The authors provide this software \"as is\", without warranty of any kind.\n" +
        "     See the LICENSE file for full terms.\n" +
        "\n" +
        "To accept, run:  graph-cli consent grant\n" +
        "To revoke,  run: graph-cli consent revoke";

    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".graph-cli");
    private static readonly string ConsentPath = Path.Combine(ConfigDir, "consent.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static bool IsConsented()
    {
        var record = Load();
        return record is { Consented: true } && record.Version == CurrentVersion;
    }

    public static ConsentStatus GetStatus()
    {
        var record = Load();
        if (record is null)
            return new ConsentStatus { Consented = false, Version = CurrentVersion };

        return new ConsentStatus
        {
            Consented = record.Consented && record.Version == CurrentVersion,
            ConsentedAt = record.ConsentedAt,
            Version = record.Version,
            RequiredVersion = CurrentVersion
        };
    }

    public static ConsentStatus Grant()
    {
        Directory.CreateDirectory(ConfigDir);
        var record = new ConsentRecord
        {
            Consented = true,
            ConsentedAt = DateTimeOffset.UtcNow,
            Version = CurrentVersion
        };
        File.WriteAllText(ConsentPath, JsonSerializer.Serialize(record, JsonOptions));
        return new ConsentStatus
        {
            Consented = true,
            ConsentedAt = record.ConsentedAt,
            Version = record.Version,
            RequiredVersion = CurrentVersion
        };
    }

    public static void Revoke()
    {
        if (File.Exists(ConsentPath))
            File.Delete(ConsentPath);
    }

    public static void EnsureConsented()
    {
        if (!IsConsented())
            throw new ConsentRequiredException();
    }

    private static ConsentRecord? Load()
    {
        if (!File.Exists(ConsentPath)) return null;
        try
        {
            return JsonSerializer.Deserialize<ConsentRecord>(File.ReadAllText(ConsentPath), JsonOptions);
        }
        catch
        {
            return null;
        }
    }
}

public class ConsentRecord
{
    public bool Consented { get; set; }
    public DateTimeOffset? ConsentedAt { get; set; }
    public string Version { get; set; } = "";
}

public class ConsentStatus
{
    public bool Consented { get; set; }
    public DateTimeOffset? ConsentedAt { get; set; }
    public string? Version { get; set; }
    public string? RequiredVersion { get; set; }
}

public class ConsentRequiredException : Exception
{
    public ConsentRequiredException()
        : base("Consent required. To use graph-cli you must first accept the terms by running 'graph-cli consent grant'. Run 'graph-cli consent show' to view the terms.") { }
}
