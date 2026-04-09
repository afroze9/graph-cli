using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class PresenceTools
{
    [McpServerTool(Name = "presence_me"), Description("Get own presence/availability status")]
    public static async Task<string> Me()
    {
        try
        {
            var result = await PresenceService.GetMeAsync();
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "presence_get"), Description("Get a user's presence/availability status")]
    public static async Task<string> Get(
        [Description("User ID")] string userId)
    {
        try
        {
            var result = await PresenceService.GetAsync(userId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "presence_batch"), Description("Get presence/availability for multiple users at once")]
    public static async Task<string> Batch(
        [Description("Comma-separated user IDs")] string userIds)
    {
        try
        {
            var result = await PresenceService.BatchAsync(userIds);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
