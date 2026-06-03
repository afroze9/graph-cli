using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class FilesTools
{
    [McpServerTool(Name = "files_list"), Description("List files and folders in a OneDrive or SharePoint drive")]
    public static async Task<string> List(
        [Description("Folder item ID (default: root)")] string? folder = null,
        [Description("Drive ID (default: current user's OneDrive)")] string? driveId = null,
        [Description("SharePoint site ID or hostname")] string? site = null,
        [Description("Number of items to retrieve (default: 25)")] int top = 25)
    {
        try
        {
            var result = await FileService.ListAsync(folder, driveId, site, top);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "files_get"), Description("Get file or folder metadata by item ID or sharing URL")]
    public static async Task<string> Get(
        [Description("Item ID or sharing URL (https://...sharepoint.com/...)")] string item,
        [Description("Drive ID (default: current user's OneDrive)")] string? driveId = null,
        [Description("SharePoint site ID or hostname")] string? site = null)
    {
        try
        {
            var result = await FileService.GetAsync(item, driveId, site);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "files_search"), Description("Search for files across OneDrive or SharePoint")]
    public static async Task<string> Search(
        [Description("Search query text")] string query,
        [Description("Drive ID (default: current user's OneDrive)")] string? driveId = null,
        [Description("SharePoint site ID or hostname")] string? site = null,
        [Description("Number of results (default: 25)")] int top = 25,
        [Description("Bypass cache and search via API")] bool refresh = false)
    {
        try
        {
            var result = await FileService.SearchAsync(query, driveId, site, top, refresh);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "files_download"), Description("Download a file from OneDrive or SharePoint to the local machine. Returns the local file path and size on success.")]
    public static async Task<string> Download(
        [Description("Item ID or sharing URL (https://...sharepoint.com/...)")] string item,
        [Description("Local output file path (default: original filename in current directory)")] string? outPath = null,
        [Description("Drive ID (default: current user's OneDrive)")] string? driveId = null,
        [Description("SharePoint site ID or hostname")] string? site = null)
    {
        try
        {
            var result = await FileService.DownloadAsync(item, outPath, driveId, site);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "files_share"), Description("Share a file or folder with others. Recipients must be in the allowed contacts list (use contacts_list to check, or ask the user to run 'graph-cli contacts allow' to add them).")]
    public static async Task<string> Share(
        [Description("Item ID or sharing URL")] string item,
        [Description("Comma-separated email addresses to share with")] string recipients,
        [Description("Permission role: read, write, or owner (default: read)")] string role = "read",
        [Description("Optional message to include in the sharing notification")] string? message = null,
        [Description("Drive ID (default: current user's OneDrive)")] string? driveId = null,
        [Description("SharePoint site ID or hostname")] string? site = null)
    {
        var emails = recipients.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (emails.Length == 0)
            return McpGraphHelper.Error("invalid_argument", "At least one recipient email is required.");

        if (!AllowedContactsService.CheckAllAndPrompt(emails, "share", interactive: false))
            return McpGraphHelper.Error("not_allowed", "One or more recipients are not in the allowed contacts list. Ask the user to run 'graph-cli contacts allow <email> --actions share' to add them.");

        try
        {
            var result = await FileService.ShareAsync(item, recipients, role, message, driveId, site);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
