using System.ComponentModel;
using GraphCli.Services;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class ContactTools
{
    [McpServerTool(Name = "contacts_list"), Description("List allowed contacts and their permitted actions (email, chat, calendar, share)")]
    public static Task<string> List(
        [Description("Filter by type: user or group")] string? type = null)
    {
        try
        {
            var result = ContactService.ListContacts(type);
            return Task.FromResult(McpGraphHelper.ToJson(result));
        }
        catch (Exception ex) { return Task.FromResult(McpGraphHelper.HandleException(ex)); }
    }

    [McpServerTool(Name = "contacts_allow"), Description("Add or update an allowed contact. This controls who can be emailed, chatted with, etc.")]
    public static Task<string> Allow(
        [Description("Email address or group identifier")] string identifier,
        [Description("Comma-separated allowed actions: email, chat, calendar, share")] string actions,
        [Description("Display name")] string? name = null,
        [Description("Contact type: user or group (default: user)")] string type = "user")
    {
        try
        {
            var result = ContactService.AllowContact(identifier, name, type, actions);
            return Task.FromResult(McpGraphHelper.ToJson(result));
        }
        catch (Exception ex) { return Task.FromResult(McpGraphHelper.HandleException(ex)); }
    }

    [McpServerTool(Name = "contacts_remove"), Description("Remove a contact from the allowed list")]
    public static Task<string> Remove(
        [Description("Email address or group identifier to remove")] string identifier)
    {
        try
        {
            var result = ContactService.RemoveContact(identifier);
            return Task.FromResult(McpGraphHelper.ToJson(result));
        }
        catch (Exception ex) { return Task.FromResult(McpGraphHelper.HandleException(ex)); }
    }
}
