using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class PagesCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var pagesCommand = new Command("pages", "SharePoint site page operations");

        pagesCommand.Subcommands.Add(BuildList(formatOption));
        pagesCommand.Subcommands.Add(BuildGet(formatOption));

        return pagesCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname (e.g. contoso.sharepoint.com:/sites/team)", Required = true };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of results to return" };
        var searchOption = new Option<string?>("--search") { Description = "Search pages by name or title (client-side, fetches all pages)" };
        var cmd = new Command("list", "List pages on a SharePoint site") { siteOption, topOption, searchOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var site = parseResult.GetValue(siteOption)!;
            var top = parseResult.GetValue(topOption);
            var search = parseResult.GetValue(searchOption);
            try
            {
                var result = await PageService.ListAsync(site, top, search);
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
        var pageArg = new Argument<string>("page-id") { Description = "Page ID" };
        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname", Required = true };
        var expandContentOption = new Option<bool>("--expand-content") { Description = "Include full page canvas layout content" };
        var cmd = new Command("get", "Get page details and optionally its content") { pageArg, siteOption, expandContentOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var pageId = parseResult.GetValue(pageArg)!;
            var site = parseResult.GetValue(siteOption)!;
            var expandContent = parseResult.GetValue(expandContentOption);
            try
            {
                var result = await PageService.GetAsync(site, pageId, expandContent);
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
