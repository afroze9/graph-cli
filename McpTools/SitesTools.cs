using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class SitesTools
{
    [McpServerTool(Name = "sites_search"), Description("Search for SharePoint sites by keyword")]
    public static async Task<string> Search(
        [Description("Search keywords to find sites")] string query,
        [Description("Number of results (default: 25)")] int top = 25,
        [Description("Skip cache and search via API")] bool refresh = false)
    {
        try
        {
            var result = await SiteService.SearchAsync(query, top, refresh);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
