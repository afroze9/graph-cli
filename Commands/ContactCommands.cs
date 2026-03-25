using System.CommandLine;
using GraphCli.Services;

namespace GraphCli.Commands;

public static class ContactCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var cmd = new Command("contacts", "Manage allowed contacts list");

        cmd.Subcommands.Add(BuildList(formatOption));
        cmd.Subcommands.Add(BuildAllow());
        cmd.Subcommands.Add(BuildRemove());

        return cmd;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var typeOption = new Option<string?>("--type") { Description = "Filter by type: user or group" };
        var cmd = new Command("list", "List allowed contacts") { typeOption };
        cmd.SetAction((parseResult, _) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var type = parseResult.GetValue(typeOption);
            var list = AllowedContactsService.Load();

            var contacts = list.Contacts.AsEnumerable();
            if (!string.IsNullOrEmpty(type))
                contacts = contacts.Where(c => c.Type.Equals(type, StringComparison.OrdinalIgnoreCase));

            var results = contacts.Select(c => new
            {
                c.Identifier,
                c.DisplayName,
                c.Type,
                AllowedActions = string.Join(", ", c.AllowedActions)
            }).ToList();

            OutputService.Print(results, format);
            return Task.CompletedTask;
        });
        return cmd;
    }

    private static Command BuildAllow()
    {
        var identifierArg = new Argument<string>("identifier") { Description = "Email address or group identifier" };
        var nameOption = new Option<string?>("--name") { Description = "Display name" };
        var typeOption = new Option<string>("--type") { DefaultValueFactory = _ => "user", Description = "Contact type: user or group" };
        var actionsOption = new Option<string>("--actions") { Description = "Comma-separated allowed actions: email, chat, calendar", Required = true };
        var cmd = new Command("allow", "Add or update an allowed contact") { identifierArg, nameOption, typeOption, actionsOption };
        cmd.SetAction((parseResult, _) =>
        {
            var identifier = parseResult.GetValue(identifierArg)!.ToLowerInvariant();
            var name = parseResult.GetValue(nameOption);
            var type = parseResult.GetValue(typeOption)!;
            var actions = parseResult.GetValue(actionsOption)!
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
                existing.AllowedActions = actions;
            }
            else
            {
                list.Contacts.Add(new AllowedContact
                {
                    Identifier = identifier,
                    DisplayName = name ?? identifier,
                    Type = type,
                    AllowedActions = actions
                });
            }

            AllowedContactsService.Save(list);
            OutputService.Print(new { status = "allowed", identifier, type, actions = string.Join(", ", actions) });
            return Task.CompletedTask;
        });
        return cmd;
    }

    private static Command BuildRemove()
    {
        var identifierArg = new Argument<string>("identifier") { Description = "Email address or group identifier to remove" };
        var cmd = new Command("remove", "Remove a contact from allowed list") { identifierArg };
        cmd.SetAction((parseResult, _) =>
        {
            var identifier = parseResult.GetValue(identifierArg)!;
            var list = AllowedContactsService.Load();
            var contact = list.FindContact(identifier);

            if (contact == null)
            {
                OutputService.PrintError("not_found", $"Contact '{identifier}' not found in allowed list.");
                Environment.ExitCode = 1;
                return Task.CompletedTask;
            }

            list.Contacts.Remove(contact);
            AllowedContactsService.Save(list);
            OutputService.Print(new { status = "removed", identifier });
            return Task.CompletedTask;
        });
        return cmd;
    }
}
