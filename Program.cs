using System.CommandLine;
using GraphCli;
using GraphCli.Commands;
using GraphCli.Services;

if (!ConsentGate.Allows(args))
{
    Console.Error.WriteLine(
        "graph-cli: consent required before this command can run.\n" +
        "\n" +
        "To use this tool, you must first accept the terms of use. Any changes made by\n" +
        "graph-cli (or by an AI assistant via its MCP server) are solely your\n" +
        "responsibility.\n" +
        "\n" +
        "  View the terms : graph-cli consent show\n" +
        "  Accept         : graph-cli consent grant\n");
    return 1;
}

var rootCommand = new RootCommand("Microsoft Graph CLI - manage mail, calendar, chat, tasks, and more");
rootCommand.Options.Add(GlobalOptions.Format);
rootCommand.Options.Add(GlobalOptions.TimeZone);

rootCommand.Subcommands.Add(ConsentCommands.Build());
rootCommand.Subcommands.Add(AuthCommands.Build());
rootCommand.Subcommands.Add(UserCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(MailCommands.Build(GlobalOptions.Format, GlobalOptions.TimeZone));
rootCommand.Subcommands.Add(CalendarCommands.Build(GlobalOptions.Format, GlobalOptions.TimeZone));
rootCommand.Subcommands.Add(ChatCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(PresenceCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(TaskCommands.Build(GlobalOptions.Format, GlobalOptions.TimeZone));
rootCommand.Subcommands.Add(ContactCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(FilesCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(PagesCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(SitesCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(ListsCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(McpCommand.Build());

return await rootCommand.Parse(args).InvokeAsync();

internal static class ConsentGate
{
    public static bool Allows(string[] args)
    {
        if (ConsentService.IsConsented()) return true;

        // No args → root help prints; let it through so new users discover `consent`.
        if (args.Length == 0) return true;

        foreach (var arg in args)
        {
            if (arg is "--help" or "-h" or "-?" or "--version") return true;
        }

        // First positional (skipping leading global options) determines the subcommand.
        var first = args.FirstOrDefault(a => !a.StartsWith('-'));
        return first is "consent";
    }
}
