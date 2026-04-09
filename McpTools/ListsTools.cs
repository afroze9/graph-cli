using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class ListsTools
{
    [McpServerTool(Name = "lists_list"), Description("List all lists on a SharePoint site")]
    public static async Task<string> List(
        [Description("SharePoint site (name, ID, or hostname path)")] string site)
    {
        try
        {
            var result = await ListService.ListAsync(site);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "lists_items"), Description("List items in a SharePoint list")]
    public static async Task<string> Items(
        [Description("SharePoint site (name, ID, or hostname path)")] string site,
        [Description("List ID or name")] string listId,
        [Description("Number of items to retrieve (default: 50)")] int top = 50,
        [Description("Comma-separated field names to select (e.g. Title,Status,Priority)")] string? fields = null,
        [Description("OData filter expression (e.g. \"fields/Status eq 'Active'\")")] string? filter = null)
    {
        try
        {
            var result = await ListService.ItemsAsync(site, listId, top, fields, filter);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
