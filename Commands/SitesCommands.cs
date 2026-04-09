using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class SitesCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var sitesCommand = new Command("sites", "SharePoint site operations");

        sitesCommand.Subcommands.Add(BuildSearch(formatOption));

        return sitesCommand;
    }

    private static Command BuildSearch(Option<string> formatOption)
    {
        var queryArg = new Argument<string>("query") { Description = "Search keywords to find sites" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of results to retrieve" };
        var refreshOption = new Option<bool>("--refresh") { DefaultValueFactory = _ => false, Description = "Skip cache and search via API" };
        var cmd = new Command("search", "Search for SharePoint sites by keyword") { queryArg, topOption, refreshOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var query = parseResult.GetValue(queryArg)!;
            var top = parseResult.GetValue(topOption);
            var refresh = parseResult.GetValue(refreshOption);
            try
            {
                var result = await SiteService.SearchAsync(query, top, refresh);
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
