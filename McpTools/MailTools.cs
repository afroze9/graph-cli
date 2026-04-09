using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class MailTools
{
    [McpServerTool(Name = "mail_list"), Description("List email messages from a mail folder")]
    public static async Task<string> List(
        [Description("Mail folder name (default: Inbox)")] string? folder = null,
        [Description("Number of messages to retrieve (default: 10)")] int top = 10,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var tz = TimeZoneService.ResolveTimeZoneId(timezone);
            var result = await MailService.ListAsync(folder, top, tz);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_get"), Description("Get full details of a specific email message")]
    public static async Task<string> Get(
        [Description("The message ID")] string messageId,
        [Description("IANA timezone")] string? timezone = null)
    {
        try
        {
            var tz = TimeZoneService.ResolveTimeZoneId(timezone);
            var result = await MailService.GetAsync(messageId, tz);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_search"), Description("Search email messages")]
    public static async Task<string> Search(
        [Description("Search query text")] string query,
        [Description("Number of results (default: 10)")] int top = 10,
        [Description("IANA timezone")] string? timezone = null)
    {
        try
        {
            var tz = TimeZoneService.ResolveTimeZoneId(timezone);
            var result = await MailService.SearchAsync(query, top, tz);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_send"), Description("Send an email. Recipients must be in the allowed contacts list (use contacts_list to check, or ask the user to run 'graph-cli contacts allow' to add them).")]
    public static async Task<string> Send(
        [Description("Comma-separated recipient email addresses")] string to,
        [Description("Email subject")] string subject,
        [Description("Email body content")] string body,
        [Description("Comma-separated CC email addresses")] string? cc = null,
        [Description("Body content type: text or html (default: text)")] string contentType = "text")
    {
        var allRecipients = to.Split(',').Select(e => e.Trim()).ToList();
        if (!string.IsNullOrEmpty(cc))
            allRecipients.AddRange(cc.Split(',').Select(e => e.Trim()));

        if (!AllowedContactsService.CheckAllAndPrompt(allRecipients, "email", interactive: false))
            return McpGraphHelper.Error("not_allowed", "One or more recipients are not in the allowed contacts list. Ask the user to run 'graph-cli contacts allow <email> --actions email' to add them.");

        try
        {
            var result = await MailService.SendAsync(to, subject, body, cc, contentType);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_draft"), Description("Create a draft email. Recipients must be in the allowed contacts list.")]
    public static async Task<string> Draft(
        [Description("Comma-separated recipient email addresses")] string to,
        [Description("Email subject")] string subject,
        [Description("Email body content")] string body,
        [Description("Body content type: text or html (default: text)")] string contentType = "text")
    {
        var recipients = to.Split(',').Select(e => e.Trim());
        if (!AllowedContactsService.CheckAllAndPrompt(recipients, "email", interactive: false))
            return McpGraphHelper.Error("not_allowed", "One or more recipients are not in the allowed contacts list.");

        try
        {
            var result = await MailService.DraftAsync(to, subject, body, contentType);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_send_draft"), Description("Send an existing draft email")]
    public static async Task<string> SendDraft(
        [Description("Draft message ID")] string messageId)
    {
        try
        {
            var result = await MailService.SendDraftAsync(messageId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_move"), Description("Move messages to a mail folder")]
    public static async Task<string> Move(
        [Description("Comma-separated message IDs")] string messageIds,
        [Description("Destination folder ID or well-known name (e.g. Inbox, Archive, DeletedItems)")] string folder)
    {
        try
        {
            var ids = messageIds.Split(',').Select(id => id.Trim()).ToArray();
            var result = await MailService.MoveAsync(ids, folder);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_delete"), Description("Delete email messages")]
    public static async Task<string> Delete(
        [Description("Comma-separated message IDs")] string messageIds)
    {
        try
        {
            var ids = messageIds.Split(',').Select(id => id.Trim()).ToArray();
            var result = await MailService.DeleteAsync(ids);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_mark_read"), Description("Mark messages as read or unread")]
    public static async Task<string> MarkRead(
        [Description("Comma-separated message IDs")] string messageIds,
        [Description("Set to true to mark as unread instead of read")] bool unread = false)
    {
        try
        {
            var ids = messageIds.Split(',').Select(id => id.Trim()).ToArray();
            var result = await MailService.MarkReadAsync(ids, unread);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_folders"), Description("List mail folders")]
    public static async Task<string> Folders(
        [Description("Parent folder ID to list child folders (optional)")] string? parent = null)
    {
        try
        {
            var result = await MailService.FoldersAsync(parent);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_attachments"), Description("List attachments on an email message")]
    public static async Task<string> Attachments(
        [Description("Message ID")] string messageId)
    {
        try
        {
            var result = await MailService.AttachmentsAsync(messageId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
