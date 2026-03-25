using System.Text.Json;
using System.Text.Json.Serialization;

namespace GraphCli.Services;

public class AllowedContactsService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".graph-cli");
    private static readonly string ContactsPath = Path.Combine(ConfigDir, "allowed-contacts.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static AllowedContactsList Load()
    {
        if (!File.Exists(ContactsPath))
            return new AllowedContactsList();

        var json = File.ReadAllText(ContactsPath);
        return JsonSerializer.Deserialize<AllowedContactsList>(json, JsonOptions) ?? new AllowedContactsList();
    }

    public static void Save(AllowedContactsList list)
    {
        Directory.CreateDirectory(ConfigDir);
        File.WriteAllText(ContactsPath, JsonSerializer.Serialize(list, JsonOptions));
    }

    /// <summary>
    /// Checks if a contact is allowed to perform the given action.
    /// If not found or not allowed, prompts the user interactively.
    /// Returns true if the action is allowed, false if denied.
    /// </summary>
    public static bool CheckAndPrompt(string identifier, string action, bool interactive = true)
    {
        var list = Load();
        var contact = list.FindContact(identifier);

        if (contact != null && contact.AllowedActions.Contains(action, StringComparer.OrdinalIgnoreCase))
            return true;

        if (!interactive || Console.IsInputRedirected)
        {
            Console.Error.WriteLine($"Contact '{identifier}' is not allowed for action '{action}'. " +
                "Run 'graph-cli contacts allow' to add them.");
            return false;
        }

        // Interactive prompt
        var contactType = contact?.Type ?? "user";
        var displayName = contact?.DisplayName ?? identifier;

        Console.Error.WriteLine();
        Console.Error.WriteLine($"  Contact '{identifier}' is not allowed for '{action}'.");
        Console.Error.WriteLine();
        Console.Error.Write($"  Allow {action} for '{identifier}'? [y/N/a] (y=yes once, a=allow and save): ");

        var response = Console.ReadLine()?.Trim().ToLowerInvariant();

        if (response == "y")
            return true;

        if (response == "a")
        {
            if (contact != null)
            {
                contact.AllowedActions.Add(action);
            }
            else
            {
                Console.Error.Write($"  Display name [{identifier}]: ");
                var name = Console.ReadLine()?.Trim();
                if (string.IsNullOrEmpty(name)) name = identifier;

                Console.Error.Write("  Type (user/group) [user]: ");
                var typeInput = Console.ReadLine()?.Trim().ToLowerInvariant();
                if (string.IsNullOrEmpty(typeInput)) typeInput = "user";

                contact = new AllowedContact
                {
                    Identifier = identifier.ToLowerInvariant(),
                    DisplayName = name,
                    Type = typeInput,
                    AllowedActions = [action]
                };
                list.Contacts.Add(contact);
            }

            Save(list);
            Console.Error.WriteLine($"  Saved. '{identifier}' is now allowed for '{action}'.");
            return true;
        }

        Console.Error.WriteLine("  Denied.");
        return false;
    }

    /// <summary>
    /// Check multiple email addresses at once (e.g., To + CC lists).
    /// Returns true only if ALL are allowed.
    /// </summary>
    public static bool CheckAllAndPrompt(IEnumerable<string> identifiers, string action, bool interactive = true)
    {
        foreach (var id in identifiers)
        {
            if (!CheckAndPrompt(id.Trim(), action, interactive))
                return false;
        }
        return true;
    }
}

public class AllowedContactsList
{
    public List<AllowedContact> Contacts { get; set; } = [];

    public AllowedContact? FindContact(string identifier)
    {
        return Contacts.FirstOrDefault(c =>
            c.Identifier.Equals(identifier, StringComparison.OrdinalIgnoreCase));
    }
}

public class AllowedContact
{
    public string Identifier { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string Type { get; set; } = "user"; // "user" or "group"
    public List<string> AllowedActions { get; set; } = []; // "email", "chat", "calendar"
}
