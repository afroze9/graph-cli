using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class ListsCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var listsCommand = new Command("lists", "SharePoint list operations");

        listsCommand.Subcommands.Add(BuildList(formatOption));
        listsCommand.Subcommands.Add(BuildItems(formatOption));

        return listsCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var siteOption = new Option<string>("--site") { Description = "SharePoint site (name, ID, or hostname path)", Required = true };
        var cmd = new Command("list", "List all lists on a SharePoint site") { siteOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var site = parseResult.GetValue(siteOption)!;
            try
            {
                var result = await ListService.ListAsync(site);
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

    private static Command BuildItems(Option<string> formatOption)
    {
        var listArg = new Argument<string>("list-id") { Description = "List ID or name" };
        var siteOption = new Option<string>("--site") { Description = "SharePoint site (name, ID, or hostname path)", Required = true };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 50, Description = "Number of items to retrieve" };
        var fieldsOption = new Option<string?>("--fields") { Description = "Comma-separated field names to select (e.g. Title,Status,Priority)" };
        var filterOption = new Option<string?>("--filter") { Description = "OData filter expression (e.g. \"fields/Status eq 'Active'\")" };
        var expandLookupsOption = new Option<string?>("--expand-lookups") { Description = "Comma-separated lookup column names to resolve (e.g. GDCProjectManager,GDCPortfolioLead). Returns {LookupId, LookupValue} alongside the raw *LookupId. Use with --fields to keep other columns visible." };
        var cmd = new Command("items", "List items in a SharePoint list") { listArg, siteOption, topOption, fieldsOption, filterOption, expandLookupsOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var listId = parseResult.GetValue(listArg)!;
            var site = parseResult.GetValue(siteOption)!;
            var top = parseResult.GetValue(topOption);
            var fields = parseResult.GetValue(fieldsOption);
            var filter = parseResult.GetValue(filterOption);
            var expandLookups = parseResult.GetValue(expandLookupsOption);
            try
            {
                var result = await ListService.ItemsAsync(site, listId, top, fields, filter, expandLookups);
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
