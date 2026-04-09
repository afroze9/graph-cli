using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class UserTools
{
    [McpServerTool(Name = "user_me"), Description("Get own user profile")]
    public static async Task<string> Me()
    {
        try
        {
            var result = await UserService.GetMeAsync();
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "user_get"), Description("Get a user by ID or email address")]
    public static async Task<string> Get(
        [Description("User ID or email address")] string userId)
    {
        try
        {
            var result = await UserService.GetUserAsync(userId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "user_search"), Description("Search users in the organization directory")]
    public static async Task<string> Search(
        [Description("Search text (matches display name or email)")] string query)
    {
        try
        {
            var result = await UserService.SearchAsync(query);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "user_manager"), Description("Get own manager")]
    public static async Task<string> Manager()
    {
        try
        {
            var result = await UserService.GetManagerAsync();
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "user_reports"), Description("Get direct reports")]
    public static async Task<string> Reports()
    {
        try
        {
            var result = await UserService.GetReportsAsync();
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
