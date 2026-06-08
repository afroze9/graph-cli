using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class TeamsCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var teamsCommand = new Command("teams", "Microsoft Teams channel operations");

        teamsCommand.Subcommands.Add(BuildListTeams(formatOption));
        teamsCommand.Subcommands.Add(BuildListChannels(formatOption));
        teamsCommand.Subcommands.Add(BuildListMessages(formatOption));
        teamsCommand.Subcommands.Add(BuildSend(formatOption));
        teamsCommand.Subcommands.Add(BuildReply(formatOption));

        return teamsCommand;
    }

    private static Command BuildListTeams(Option<string> formatOption)
    {
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of teams to retrieve" };
        var cmd = new Command("list", "List Teams that the current user has joined") { topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var top = parseResult.GetValue(topOption);
            try
            {
                var result = await TeamsService.ListTeamsAsync(top);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildListChannels(Option<string> formatOption)
    {
        var teamIdArg = new Argument<string>("team-id") { Description = "Team ID" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of channels to retrieve" };
        var cmd = new Command("channels", "List channels in a team") { teamIdArg, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var teamId = parseResult.GetValue(teamIdArg)!;
            var top = parseResult.GetValue(topOption);
            try
            {
                var result = await TeamsService.ListChannelsAsync(teamId, top);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildListMessages(Option<string> formatOption)
    {
        var teamIdArg = new Argument<string>("team-id") { Description = "Team ID" };
        var channelIdArg = new Argument<string>("channel-id") { Description = "Channel ID" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 20, Description = "Number of messages to retrieve" };
        var cmd = new Command("messages", "List messages in a Teams channel") { teamIdArg, channelIdArg, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var teamId = parseResult.GetValue(teamIdArg)!;
            var channelId = parseResult.GetValue(channelIdArg)!;
            var top = parseResult.GetValue(topOption);
            try
            {
                var result = await TeamsService.ListMessagesAsync(teamId, channelId, top);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildSend(Option<string> formatOption)
    {
        var teamIdArg = new Argument<string>("team-id") { Description = "Team ID" };
        var channelIdArg = new Argument<string>("channel-id") { Description = "Channel ID" };
        var messageOption = new Option<string>("--message") { Description = "Message text", Required = true };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Content type: text or html" };
        var mentionsOption = new Option<string?>("--mentions") { Description = "Comma-separated user emails or AAD IDs to @-mention (requires --content-type html)" };
        var cmd = new Command("send", "Send a message to a Teams channel") { teamIdArg, channelIdArg, messageOption, contentTypeOption, mentionsOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var teamId = parseResult.GetValue(teamIdArg)!;
            var channelId = parseResult.GetValue(channelIdArg)!;
            var message = parseResult.GetValue(messageOption)!;
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";
            var mentionsCsv = parseResult.GetValue(mentionsOption);
            try
            {
                var mentions = string.IsNullOrEmpty(mentionsCsv)
                    ? null
                    : mentionsCsv.Split(',').Select(m => m.Trim()).ToArray();
                var result = await TeamsService.SendMessageAsync(teamId, channelId, message, contentType, mentions);
                OutputService.Print(result);
            }
            catch (ArgumentException ex)
            {
                OutputService.PrintError("invalid_argument", ex.Message);
                Environment.ExitCode = 1;
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildReply(Option<string> formatOption)
    {
        var teamIdArg = new Argument<string>("team-id") { Description = "Team ID" };
        var channelIdArg = new Argument<string>("channel-id") { Description = "Channel ID" };
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID of the thread root to reply to" };
        var messageOption = new Option<string>("--message") { Description = "Reply text", Required = true };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Content type: text or html" };
        var mentionsOption = new Option<string?>("--mentions") { Description = "Comma-separated user emails or AAD IDs to @-mention (requires --content-type html)" };
        var cmd = new Command("reply", "Reply to a message thread in a Teams channel") { teamIdArg, channelIdArg, messageIdArg, messageOption, contentTypeOption, mentionsOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var teamId = parseResult.GetValue(teamIdArg)!;
            var channelId = parseResult.GetValue(channelIdArg)!;
            var messageId = parseResult.GetValue(messageIdArg)!;
            var message = parseResult.GetValue(messageOption)!;
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";
            var mentionsCsv = parseResult.GetValue(mentionsOption);
            try
            {
                var mentions = string.IsNullOrEmpty(mentionsCsv)
                    ? null
                    : mentionsCsv.Split(',').Select(m => m.Trim()).ToArray();
                var result = await TeamsService.ReplyAsync(teamId, channelId, messageId, message, contentType, mentions);
                OutputService.Print(result);
            }
            catch (ArgumentException ex)
            {
                OutputService.PrintError("invalid_argument", ex.Message);
                Environment.ExitCode = 1;
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }
}
