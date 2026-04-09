using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class PresenceCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var presenceCommand = new Command("presence", "Presence/status operations");

        presenceCommand.Subcommands.Add(BuildMe(formatOption));
        presenceCommand.Subcommands.Add(BuildGet(formatOption));
        presenceCommand.Subcommands.Add(BuildBatch(formatOption));

        return presenceCommand;
    }

    private static Command BuildMe(Option<string> formatOption)
    {
        var cmd = new Command("me", "Get own presence status");
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PresenceService.GetMeAsync();
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

    private static Command BuildGet(Option<string> formatOption)
    {
        var userIdArg = new Argument<string>("user-id") { Description = "User ID" };
        var cmd = new Command("get", "Get a user's presence") { userIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var userId = parseResult.GetValue(userIdArg)!;
            try
            {
                var result = await PresenceService.GetAsync(userId);
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

    private static Command BuildBatch(Option<string> formatOption)
    {
        var userIdsOption = new Option<string>("--user-ids") { Description = "Comma-separated user IDs", Required = true };
        var cmd = new Command("batch", "Get presence for multiple users") { userIdsOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var userIds = parseResult.GetValue(userIdsOption)!;
            try
            {
                var result = await PresenceService.BatchAsync(userIds);
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
}
