using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class FilesCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var filesCommand = new Command("files", "OneDrive and SharePoint file operations");

        filesCommand.Subcommands.Add(BuildList(formatOption));
        filesCommand.Subcommands.Add(BuildGet(formatOption));
        filesCommand.Subcommands.Add(BuildDownload());
        filesCommand.Subcommands.Add(BuildSearch(formatOption));
        filesCommand.Subcommands.Add(BuildShare(formatOption));

        return filesCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var folderOption = new Option<string?>("--folder") { Description = "Folder item ID (default: root)" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname (e.g. contoso.sharepoint.com:/sites/team)" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of items to retrieve" };
        var cmd = new Command("list", "List files and folders in a drive") { folderOption, driveIdOption, siteOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var folder = parseResult.GetValue(folderOption);
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);
            var top = parseResult.GetValue(topOption);
            try
            {
                var result = await FileService.ListAsync(folder, driveId, site, top);
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
        var itemArg = new Argument<string>("item") { Description = "Item ID, or a sharing URL (https://...sharepoint.com/...)" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname" };
        var cmd = new Command("get", "Get file or folder metadata") { itemArg, driveIdOption, siteOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var item = parseResult.GetValue(itemArg)!;
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);
            try
            {
                var result = await FileService.GetAsync(item, driveId, site);
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

    private static Command BuildDownload()
    {
        var itemArg = new Argument<string>("item") { Description = "Item ID, or a sharing URL (https://...sharepoint.com/...)" };
        var outOption = new Option<string?>("--out") { Description = "Output file path (default: original filename in current directory)" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname" };
        var cmd = new Command("download", "Download a file") { itemArg, outOption, driveIdOption, siteOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var item = parseResult.GetValue(itemArg)!;
            var outPath = parseResult.GetValue(outOption);
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);
            try
            {
                var result = await FileService.DownloadAsync(item, outPath, driveId, site);
                OutputService.Print(result);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildSearch(Option<string> formatOption)
    {
        var queryArg = new Argument<string>("query") { Description = "Search query text" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of results to retrieve" };
        var refreshOption = new Option<bool>("--refresh") { DefaultValueFactory = _ => false, Description = "Bypass cache and search via API" };
        var cmd = new Command("search", "Search for files across OneDrive or SharePoint") { queryArg, driveIdOption, siteOption, topOption, refreshOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var query = parseResult.GetValue(queryArg)!;
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);
            var top = parseResult.GetValue(topOption);
            var refresh = parseResult.GetValue(refreshOption);
            try
            {
                var result = await FileService.SearchAsync(query, driveId, site, top, refresh);
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

    private static Command BuildShare(Option<string> formatOption)
    {
        var itemArg = new Argument<string>("item") { Description = "Item ID, or a sharing URL (https://...sharepoint.com/...)" };
        var recipientsOption = new Option<string>("--recipients") { Description = "Comma-separated email addresses to share with", Required = true };
        var roleOption = new Option<string>("--role") { DefaultValueFactory = _ => "read", Description = "Permission role: read, write, or owner" };
        var messageOption = new Option<string?>("--message") { Description = "Optional message to include in the sharing notification" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname" };
        var cmd = new Command("share", "Share a file or folder with others") { itemArg, recipientsOption, roleOption, messageOption, driveIdOption, siteOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var item = parseResult.GetValue(itemArg)!;
            var recipients = parseResult.GetValue(recipientsOption)!;
            var role = parseResult.GetValue(roleOption) ?? "read";
            var message = parseResult.GetValue(messageOption);
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);

            var emails = recipients.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (emails.Length == 0)
            {
                OutputService.PrintError("invalid_argument", "At least one recipient email is required.");
                Environment.ExitCode = 1;
                return;
            }

            if (!AllowedContactsService.CheckAllAndPrompt(emails, "share"))
            {
                Environment.ExitCode = 1;
                return;
            }

            try
            {
                var result = await FileService.ShareAsync(item, recipients, role, message, driveId, site);
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
