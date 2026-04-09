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
            var result = ContactService.ListContacts(parseResult.GetValue(typeOption));
            OutputService.Print(result, format);
            return Task.CompletedTask;
        });
        return cmd;
    }

    private static Command BuildAllow()
    {
        var identifierArg = new Argument<string>("identifier") { Description = "Email address or group identifier" };
        var nameOption = new Option<string?>("--name") { Description = "Display name" };
        var typeOption = new Option<string>("--type") { DefaultValueFactory = _ => "user", Description = "Contact type: user or group" };
        var actionsOption = new Option<string>("--actions") { Description = "Comma-separated allowed actions: email, chat, calendar, share", Required = true };
        var cmd = new Command("allow", "Add or update an allowed contact") { identifierArg, nameOption, typeOption, actionsOption };
        cmd.SetAction((parseResult, _) =>
        {
            var result = ContactService.AllowContact(
                parseResult.GetValue(identifierArg)!,
                parseResult.GetValue(nameOption),
                parseResult.GetValue(typeOption)!,
                parseResult.GetValue(actionsOption)!);
            OutputService.Print(result);
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
            try
            {
                var result = ContactService.RemoveContact(parseResult.GetValue(identifierArg)!);
                OutputService.Print(result);
            }
            catch (InvalidOperationException ex)
            {
                OutputService.PrintError("not_found", ex.Message);
                Environment.ExitCode = 1;
            }
            return Task.CompletedTask;
        });
        return cmd;
    }
}
