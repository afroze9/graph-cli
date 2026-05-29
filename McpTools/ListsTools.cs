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
        [Description("OData filter expression (e.g. \"fields/Status eq 'Active'\")")] string? filter = null,
        [Description("Comma-separated lookup column names to resolve (e.g. GDCProjectManager,GDCPortfolioLead). Returns {LookupId, LookupValue} alongside the raw *LookupId. Use with 'fields' to keep other columns visible.")] string? expandLookups = null)
    {
        try
        {
            var result = await ListService.ItemsAsync(site, listId, top, fields, filter, expandLookups);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "lists_columns"), Description("List column definitions (displayName + internal name) for a SharePoint list, to map mangled internal field names back to their real question/column text")]
    public static async Task<string> Columns(
        [Description("SharePoint site (name, ID, or hostname path)")] string site,
        [Description("List ID or name")] string listId)
    {
        try
        {
            var result = await ListService.ColumnsAsync(site, listId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
