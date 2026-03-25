using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph;
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
                var client = await GraphClientProvider.CreateAsync();
                var me = await client.Me.GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName", "jobTitle", "department", "officeLocation", "mobilePhone", "businessPhones"];
                }, ct);
                OutputService.Print(new
                {
                    me!.Id,
                    me.DisplayName,
                    me.Mail,
                    me.UserPrincipalName,
                    me.JobTitle,
                    me.Department,
                    me.OfficeLocation,
                    me.MobilePhone,
                    me.BusinessPhones
                }, format);
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
                var client = await GraphClientProvider.CreateAsync();
                var user = await client.Users[userId].GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName", "jobTitle", "department"];
                }, ct);
                OutputService.Print(new
                {
                    user!.Id,
                    user.DisplayName,
                    user.Mail,
                    user.UserPrincipalName,
                    user.JobTitle,
                    user.Department
                }, format);
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
                var client = await GraphClientProvider.CreateAsync();
                var users = await client.Users.GetAsync(r =>
                {
                    r.QueryParameters.Filter = $"startsWith(displayName,'{query}') or startsWith(mail,'{query}')";
                    r.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName", "jobTitle"];
                    r.QueryParameters.Top = 25;
                }, ct);
                var results = users?.Value?.Select(u => new
                {
                    u.Id,
                    u.DisplayName,
                    u.Mail,
                    u.UserPrincipalName,
                    u.JobTitle
                }).ToList();
                OutputService.Print(results, format);
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
                var client = await GraphClientProvider.CreateAsync();
                var manager = await client.Me.Manager.GetAsync(cancellationToken: ct);
                var props = manager?.AdditionalData;
                OutputService.Print(new
                {
                    id = GetProp(props, "id"),
                    displayName = GetProp(props, "displayName"),
                    mail = GetProp(props, "mail"),
                    userPrincipalName = GetProp(props, "userPrincipalName"),
                    jobTitle = GetProp(props, "jobTitle")
                }, format);
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
                var client = await GraphClientProvider.CreateAsync();
                var reports = await client.Me.DirectReports.GetAsync(cancellationToken: ct);
                var results = reports?.Value?.Select(r => new
                {
                    id = GetProp(r.AdditionalData, "id"),
                    displayName = GetProp(r.AdditionalData, "displayName"),
                    mail = GetProp(r.AdditionalData, "mail"),
                    userPrincipalName = GetProp(r.AdditionalData, "userPrincipalName")
                }).ToList();
                OutputService.Print(results, format);
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
        return userCommand;
    }

    private static string? GetProp(IDictionary<string, object>? data, string key)
    {
        if (data != null && data.TryGetValue(key, out var value))
            return value?.ToString();
        return null;
    }
}
