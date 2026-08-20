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

    [McpServerTool(Name = "mail_send"), Description("Send an email with optional file attachments. Recipients must be in the allowed contacts list (use contacts_list to check, or ask the user to run 'graph-cli contacts allow' to add them). To @-mention people (puts the @ glyph on the mail in their Outlook and makes it show under \"Mentioned mail\"), set contentType=html, include <at id=\"N\">Name</at> tags in the body, and pass a comma-separated email list in `mentions` where index N matches the at-tag id. Mentioned people are added to the To line automatically.")]
    public static async Task<string> Send(
        [Description("Comma-separated recipient email addresses")] string to,
        [Description("Email subject")] string subject,
        [Description("Email body content")] string body,
        [Description("Comma-separated CC email addresses")] string? cc = null,
        [Description("Body content type: text or html (default: text). Required to be html when mentions is set.")] string contentType = "text",
        [Description("Comma-separated file paths to attach (e.g. /path/to/report.pdf,/path/to/data.xlsx)")] string? attachments = null,
        [Description("Comma-separated emails to @-mention. Body must contain <at id=\"N\">Name</at> tags (N is the zero-based index into this list).")] string? mentions = null)
    {
        var mentionList = McpGraphHelper.SplitCsv(mentions);

        var allRecipients = to.Split(',').Select(e => e.Trim()).ToList();
        if (!string.IsNullOrEmpty(cc))
            allRecipients.AddRange(cc.Split(',').Select(e => e.Trim()));
        // Mentioned people become recipients, so they pass the same gate.
        if (mentionList != null)
            allRecipients.AddRange(mentionList);

        if (!AllowedContactsService.CheckAllAndPrompt(allRecipients, "email", interactive: false))
            return McpGraphHelper.Error("not_allowed", "One or more recipients are not in the allowed contacts list. Ask the user to run 'graph-cli contacts allow <email> --actions email' to add them.");

        try
        {
            var attachmentPaths = McpGraphHelper.SplitCsv(attachments);
            var result = await MailService.SendAsync(to, subject, body, cc, contentType, attachmentPaths, mentionList);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_mentions"), Description("List the @-mentions stored on an email message. Returns isMentioned (whether the signed-in user is mentioned) and each mention's name and address. Use it to confirm a mail_send with mentions landed correctly.")]
    public static async Task<string> Mentions(
        [Description("The message ID")] string messageId)
    {
        try
        {
            var result = await MailService.MentionsAsync(messageId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_draft"), Description("Create a draft email. Recipients must be in the allowed contacts list. Supports @-mentions the same way mail_send does.")]
    public static async Task<string> Draft(
        [Description("Comma-separated recipient email addresses")] string to,
        [Description("Email subject")] string subject,
        [Description("Email body content")] string body,
        [Description("Body content type: text or html (default: text). Required to be html when mentions is set.")] string contentType = "text",
        [Description("Comma-separated emails to @-mention. Body must contain <at id=\"N\">Name</at> tags (N is the zero-based index into this list).")] string? mentions = null)
    {
        var mentionList = McpGraphHelper.SplitCsv(mentions);

        var recipients = to.Split(',').Select(e => e.Trim()).ToList();
        if (mentionList != null)
            recipients.AddRange(mentionList);

        if (!AllowedContactsService.CheckAllAndPrompt(recipients, "email", interactive: false))
            return McpGraphHelper.Error("not_allowed", "One or more recipients are not in the allowed contacts list.");

        try
        {
            var result = await MailService.DraftAsync(to, subject, body, contentType, mentionList);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_reply"), Description("Reply to an email message (preserves thread). Use replyAll=true to reply to all recipients. Any additional CC/BCC recipients must be in the allowed contacts list.")]
    public static async Task<string> Reply(
        [Description("Message ID to reply to")] string messageId,
        [Description("Reply body (prepended to quoted original)")] string body,
        [Description("Set to true to reply-all instead of sender-only (default: false)")] bool replyAll = false,
        [Description("Comma-separated additional CC emails")] string? cc = null,
        [Description("Comma-separated additional BCC emails")] string? bcc = null,
        [Description("Body content type: text (keeps quoted thread) or html (replaces body). Default: text")] string contentType = "text",
        [Description("Comma-separated file paths to attach")] string? attachments = null,
        [Description("Create a draft instead of sending (default: false)")] bool draft = false)
    {
        var added = new List<string>();
        if (!string.IsNullOrEmpty(cc)) added.AddRange(cc.Split(',').Select(e => e.Trim()));
        if (!string.IsNullOrEmpty(bcc)) added.AddRange(bcc.Split(',').Select(e => e.Trim()));
        if (added.Count > 0 && !AllowedContactsService.CheckAllAndPrompt(added, "email", interactive: false))
            return McpGraphHelper.Error("not_allowed", "One or more added recipients are not in the allowed contacts list.");

        try
        {
            var attachmentPaths = string.IsNullOrEmpty(attachments)
                ? null
                : attachments.Split(',').Select(p => p.Trim()).ToArray();
            var result = await MailService.ReplyAsync(messageId, body, contentType, cc, bcc, attachmentPaths, replyAll, draft);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "mail_forward"), Description("Forward an email message to new recipients (preserves thread). Recipients must be in the allowed contacts list.")]
    public static async Task<string> Forward(
        [Description("Message ID to forward")] string messageId,
        [Description("Comma-separated recipient email addresses")] string to,
        [Description("Forward body (prepended to quoted original)")] string body,
        [Description("Comma-separated CC emails")] string? cc = null,
        [Description("Comma-separated BCC emails")] string? bcc = null,
        [Description("Body content type: text (keeps quoted thread) or html (replaces body). Default: text")] string contentType = "text",
        [Description("Comma-separated file paths to attach")] string? attachments = null,
        [Description("Create a draft instead of sending (default: false)")] bool draft = false)
    {
        var allRecipients = to.Split(',').Select(e => e.Trim()).ToList();
        if (!string.IsNullOrEmpty(cc)) allRecipients.AddRange(cc.Split(',').Select(e => e.Trim()));
        if (!string.IsNullOrEmpty(bcc)) allRecipients.AddRange(bcc.Split(',').Select(e => e.Trim()));
        if (!AllowedContactsService.CheckAllAndPrompt(allRecipients, "email", interactive: false))
            return McpGraphHelper.Error("not_allowed", "One or more recipients are not in the allowed contacts list.");

        try
        {
            var attachmentPaths = string.IsNullOrEmpty(attachments)
                ? null
                : attachments.Split(',').Select(p => p.Trim()).ToArray();
            var result = await MailService.ForwardAsync(messageId, to, body, contentType, cc, bcc, attachmentPaths, draft);
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
