using System.CommandLine;
using GraphCli;
using GraphCli.Commands;

var rootCommand = new RootCommand("Microsoft Graph CLI - manage mail, calendar, chat, tasks, and more");
rootCommand.Options.Add(GlobalOptions.Format);
rootCommand.Options.Add(GlobalOptions.TimeZone);

rootCommand.Subcommands.Add(AuthCommands.Build());
rootCommand.Subcommands.Add(UserCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(MailCommands.Build(GlobalOptions.Format, GlobalOptions.TimeZone));
rootCommand.Subcommands.Add(CalendarCommands.Build(GlobalOptions.Format, GlobalOptions.TimeZone));
rootCommand.Subcommands.Add(ChatCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(PresenceCommands.Build(GlobalOptions.Format));
rootCommand.Subcommands.Add(TaskCommands.Build(GlobalOptions.Format, GlobalOptions.TimeZone));
rootCommand.Subcommands.Add(ContactCommands.Build(GlobalOptions.Format));

return await rootCommand.Parse(args).InvokeAsync();
