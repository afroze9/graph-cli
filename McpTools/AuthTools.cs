using System.ComponentModel;
using GraphCli.Services;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class AuthTools
{
    [McpServerTool(Name = "auth_status"), Description("Check current authentication status. If not logged in, user must run 'graph-cli auth login' in a terminal.")]
    public static async Task<string> Status()
    {
        try
        {
            var authService = new AuthService();
            var status = await authService.GetStatusAsync();
            return McpGraphHelper.ToJson(status);
        }
        catch (Exception ex)
        {
            return McpGraphHelper.Error("auth_error", ex.Message);
        }
    }
}
