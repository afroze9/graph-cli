using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class ChatTools
{
    [McpServerTool(Name = "chat_list"), Description("List recent Teams chats")]
    public static async Task<string> List(
        [Description("Number of chats to retrieve (default: 20)")] int top = 20)
    {
        try
        {
            var result = await ChatService.ListAsync(top);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_search"), Description("Search Teams chats by topic, member display name, or member email. Cache-backed: matches against the local cache first; on miss (or with refresh=true) pulls the most recent chats from Graph with members expanded and merges into the cache before re-searching.")]
    public static async Task<string> Search(
        [Description("Search text (case-insensitive match against chat topic, member display name, or member email)")] string query,
        [Description("Max results to return (default: 20)")] int top = 20,
        [Description("Skip cache, force a fresh paginated API fetch (up to maxDepth) and merge into cache")] bool refresh = false,
        [Description("Maximum number of chats to pull from the API when refreshing or on cache miss, ordered by most recently active (default: 200)")] int maxDepth = 200)
    {
        try
        {
            var result = await ChatService.SearchAsync(query, top, refresh, maxDepth);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_find_with"), Description("Find Teams chats whose members include a specific user. Uses a Graph $filter on members/userId, so it locates 1:1 and group chats with that person without scanning the full chat list.")]
    public static async Task<string> FindWith(
        [Description("User email, UPN, or AAD object ID")] string user,
        [Description("Chat type filter: oneOnOne, group, or all (default: all)")] string type = "all",
        [Description("Max results to return (default: 20)")] int top = 20)
    {
        try
        {
            var result = await ChatService.FindWithAsync(user, type, top);
            return McpGraphHelper.ToJson(result);
        }
        catch (ArgumentException ex) { return McpGraphHelper.Error("invalid_argument", ex.Message); }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_get"), Description("Get details of a specific Teams chat")]
    public static async Task<string> Get(
        [Description("Chat ID")] string chatId)
    {
        try
        {
            var result = await ChatService.GetAsync(chatId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_create"), Description("Create a new Teams chat. Members must be in the allowed contacts list.")]
    public static async Task<string> Create(
        [Description("Comma-separated member email addresses")] string members,
        [Description("Chat topic (for group chats)")] string? topic = null,
        [Description("Chat type: oneOnOne or group (default: oneOnOne)")] string type = "oneOnOne")
    {
        var memberEmails = members.Split(',').Select(e => e.Trim());
        if (!AllowedContactsService.CheckAllAndPrompt(memberEmails, "chat", interactive: false))
            return McpGraphHelper.Error("not_allowed", "One or more members are not in the allowed contacts list. Ask the user to run 'graph-cli contacts allow <email> --actions chat' to add them.");

        try
        {
            var result = await ChatService.CreateAsync(members, topic, type);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_members"), Description("List members of a Teams chat")]
    public static async Task<string> Members(
        [Description("Chat ID")] string chatId)
    {
        try
        {
            var result = await ChatService.MembersAsync(chatId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_messages"), Description("List messages in a Teams chat. Each message includes a reactions array (emoji reactions such as a 👍 thumbs-up, reported with reactionType, the reacting user's id, and the timestamp) — useful for detecting when someone acknowledged a message with a reaction rather than a text reply.")]
    public static async Task<string> Messages(
        [Description("Chat ID")] string chatId,
        [Description("Number of messages to retrieve (default: 20)")] int top = 20)
    {
        try
        {
            var result = await ChatService.MessagesAsync(chatId, top);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_send_image"), Description("Send an image file as an inline attachment in a Teams chat. The image is uploaded as hosted content so it renders inline (not as a download link). Supports PNG, JPEG, GIF, and WebP. Chat must be in the allowed contacts list.")]
    public static async Task<string> SendImage(
        [Description("Chat ID")] string chatId,
        [Description("Absolute path to the image file on disk (PNG, JPEG, GIF, or WebP)")] string imagePath,
        [Description("Optional caption text displayed above the image")] string? caption = null)
    {
        if (!AllowedContactsService.CheckAndPrompt(chatId, "chat", interactive: false))
            return McpGraphHelper.Error("not_allowed", "This chat is not in the allowed contacts list. Ask the user to run 'graph-cli contacts allow <chatId> --actions chat' to add it.");

        try
        {
            var result = await ChatService.SendImageAsync(chatId, imagePath, caption);
            return McpGraphHelper.ToJson(result);
        }
        catch (ArgumentException ex) { return McpGraphHelper.Error("invalid_argument", ex.Message); }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_send"), Description("Send a message in a Teams chat. Chat must be in the allowed contacts list. To @-mention users (fires Teams notifications), set contentType=html, include <at id=\"N\">Name</at> tags in the message body, and pass a comma-separated emails/AAD IDs list in `mentions` where index N matches the at-tag id.")]
    public static async Task<string> Send(
        [Description("Chat ID")] string chatId,
        [Description("Message text")] string message,
        [Description("Content type: text or html (default: text). Required to be html when mentions is set.")] string contentType = "text",
        [Description("Comma-separated emails or AAD user IDs to @-mention. Body must contain <at id=\"N\">Name</at> tags (N is zero-based index into this list).")] string? mentions = null)
    {
        if (!AllowedContactsService.CheckAndPrompt(chatId, "chat", interactive: false))
            return McpGraphHelper.Error("not_allowed", "This chat is not in the allowed contacts list. Ask the user to run 'graph-cli contacts allow <chatId> --actions chat' to add it.");

        try
        {
            var mentionList = string.IsNullOrEmpty(mentions)
                ? null
                : mentions.Split(',').Select(m => m.Trim()).ToArray();
            var result = await ChatService.SendAsync(chatId, message, contentType, mentionList);
            return McpGraphHelper.ToJson(result);
        }
        catch (ArgumentException ex) { return McpGraphHelper.Error("invalid_argument", ex.Message); }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "chat_reply"), Description("Reply to a message in a Teams chat. Chat must be in the allowed contacts list. Supports @-mentions the same way chat_send does.")]
    public static async Task<string> Reply(
        [Description("Chat ID")] string chatId,
        [Description("Message ID to reply to")] string messageId,
        [Description("Reply text")] string message,
        [Description("Content type: text or html (default: text). Required to be html when mentions is set.")] string contentType = "text",
        [Description("Comma-separated emails or AAD user IDs to @-mention. Body must contain <at id=\"N\">Name</at> tags (N is zero-based index into this list).")] string? mentions = null)
    {
        if (!AllowedContactsService.CheckAndPrompt(chatId, "chat", interactive: false))
            return McpGraphHelper.Error("not_allowed", "This chat is not in the allowed contacts list. Ask the user to run 'graph-cli contacts allow <chatId> --actions chat' to add it.");

        try
        {
            var mentionList = string.IsNullOrEmpty(mentions)
                ? null
                : mentions.Split(',').Select(m => m.Trim()).ToArray();
            var result = await ChatService.ReplyAsync(chatId, messageId, message, contentType, mentionList);
            return McpGraphHelper.ToJson(result);
        }
        catch (ArgumentException ex) { return McpGraphHelper.Error("invalid_argument", ex.Message); }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
