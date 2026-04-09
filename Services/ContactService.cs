namespace GraphCli.Services;

public static class ContactService
{
    public static object ListContacts(string? type)
    {
        var list = AllowedContactsService.Load();
        var contacts = list.Contacts.AsEnumerable();
        if (!string.IsNullOrEmpty(type))
            contacts = contacts.Where(c => c.Type.Equals(type, StringComparison.OrdinalIgnoreCase));

        return contacts.Select(c => new
        {
            c.Identifier,
            c.DisplayName,
            c.Type,
            AllowedActions = string.Join(", ", c.AllowedActions)
        }).ToList();
    }

    public static object AllowContact(string identifier, string? name, string type, string actions)
    {
        identifier = identifier.ToLowerInvariant();
        var actionList = actions
            .Split(',')
            .Select(a => a.Trim().ToLowerInvariant())
            .Where(a => !string.IsNullOrEmpty(a))
            .ToList();

        var list = AllowedContactsService.Load();
        var existing = list.FindContact(identifier);

        if (existing != null)
        {
            if (!string.IsNullOrEmpty(name)) existing.DisplayName = name;
            existing.Type = type;
            existing.AllowedActions = actionList;
        }
        else
        {
            list.Contacts.Add(new AllowedContact
            {
                Identifier = identifier,
                DisplayName = name ?? identifier,
                Type = type,
                AllowedActions = actionList
            });
        }

        AllowedContactsService.Save(list);
        return new { status = "allowed", identifier, type, actions = string.Join(", ", actionList) };
    }

    public static object RemoveContact(string identifier)
    {
        var list = AllowedContactsService.Load();
        var contact = list.FindContact(identifier);

        if (contact == null)
            throw new InvalidOperationException($"Contact '{identifier}' not found in allowed list.");

        list.Contacts.Remove(contact);
        AllowedContactsService.Save(list);
        return new { status = "removed", identifier };
    }
}
