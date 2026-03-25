using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph;
using Microsoft.Graph.Communications.GetPresencesByUserId;
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
                var client = await GraphClientProvider.CreateAsync();
                var presence = await client.Me.Presence.GetAsync(cancellationToken: ct);
                OutputService.Print(new
                {
                    Availability = presence!.Availability,
                    Activity = presence.Activity,
                    StatusMessage = presence.StatusMessage?.Message?.Content
                }, format);
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
                var client = await GraphClientProvider.CreateAsync();
                var presence = await client.Communications.Presences[userId].GetAsync(cancellationToken: ct);
                OutputService.Print(new
                {
                    presence!.Id,
                    Availability = presence.Availability,
                    Activity = presence.Activity
                }, format);
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
                var client = await GraphClientProvider.CreateAsync();
                var ids = userIds.Split(',').Select(id => id.Trim()).ToList();
                var presences = await client.Communications.GetPresencesByUserId.PostAsGetPresencesByUserIdPostResponseAsync(
                    new GetPresencesByUserIdPostRequestBody { Ids = ids },
                    cancellationToken: ct);
                var results = presences?.Value?.Select(p => new
                {
                    p.Id,
                    Availability = p.Availability,
                    Activity = p.Activity
                }).ToList();
                OutputService.Print(results, format);
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
