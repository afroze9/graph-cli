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
                // Check cache first
                if (!refresh)
                {
                    var cached = SiteCacheService.Search(query, top);
                    if (cached.Count > 0)
                    {
                        var cachedResults = cached.Select(s => new
                        {
                            s.Id,
                            s.Name,
                            s.DisplayName,
                            s.WebUrl,
                            Source = "cache"
                        }).ToList();
                        OutputService.Print(cachedResults, format);
                        return;
                    }
                }

                var client = await GraphClientProvider.CreateAsync();

                var sites = await client.Sites.GetAsync(r =>
                {
                    r.QueryParameters.Search = query;
                    r.QueryParameters.Top = top;
                    r.QueryParameters.Select = ["id", "name", "displayName", "webUrl", "description"];
                }, ct);

                // Cache results
                if (sites?.Value != null)
                {
                    SiteCacheService.UpsertMany(sites.Value
                        .Where(s => s.Id != null && s.Name != null)
                        .Select(s => (s.Id!, s.Name!, s.DisplayName, s.WebUrl)));
                }

                var results = sites?.Value?.Select(s => new
                {
                    s.Id,
                    s.Name,
                    s.DisplayName,
                    s.Description,
                    s.WebUrl
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
