using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class PagesTools
{
    [McpServerTool(Name = "pages_list"), Description("List pages on a SharePoint site")]
    public static async Task<string> List(
        [Description("SharePoint site ID or hostname (e.g. contoso.sharepoint.com:/sites/team)")] string site,
        [Description("Number of results (default: 25)")] int top = 25,
        [Description("Search pages by name or title")] string? search = null)
    {
        try
        {
            var result = await PageService.ListAsync(site, top, search);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_get"), Description("Get page details and optionally its content")]
    public static async Task<string> Get(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Include full page canvas layout content")] bool expandContent = false)
    {
        try
        {
            var result = await PageService.GetAsync(site, pageId, expandContent);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
