using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class UserCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var userCommand = new Command("user", "User operations");

        // user me
        var meCommand = new Command("me", "Get own profile");
        meCommand.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await UserService.GetMeAsync();
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });

        // user get <user-id-or-email>
        var userIdArg = new Argument<string>("user-id-or-email") { Description = "User ID or email address" };
        var getCommand = new Command("get", "Get user by ID or email") { userIdArg };
        getCommand.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var userId = parseResult.GetValue(userIdArg)!;
            try
            {
                var result = await UserService.GetUserAsync(userId);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });

        // user search --query <text>
        var queryOption = new Option<string>("--query") { Description = "Search text", Required = true };
        var searchCommand = new Command("search", "Search users in directory") { queryOption };
        searchCommand.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var query = parseResult.GetValue(queryOption)!;
            try
            {
                var result = await UserService.SearchAsync(query);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });

        // user manager
        var managerCommand = new Command("manager", "Get own manager");
        managerCommand.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await UserService.GetManagerAsync();
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });

        // user reports
        var reportsCommand = new Command("reports", "Get direct reports");
        reportsCommand.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await UserService.GetReportsAsync();
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });

        // user photo <user-id-or-email> --out <path>
        var photoUserIdArg = new Argument<string>("user-id-or-email") { Description = "User ID or email address" };
        var photoOutOption = new Option<string>("--out") { Description = "Output file path (e.g. avatar.jpg)", Required = true };
        var photoCommand = new Command("photo", "Download a user's profile photo") { photoUserIdArg, photoOutOption };
        photoCommand.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var userId = parseResult.GetValue(photoUserIdArg)!;
            var outPath = parseResult.GetValue(photoOutOption)!;
            try
            {
                var result = await UserService.GetPhotoAsync(userId, outPath);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });

        userCommand.Subcommands.Add(meCommand);
        userCommand.Subcommands.Add(getCommand);
        userCommand.Subcommands.Add(searchCommand);
        userCommand.Subcommands.Add(managerCommand);
        userCommand.Subcommands.Add(reportsCommand);
        userCommand.Subcommands.Add(photoCommand);
        return userCommand;
    }
}
