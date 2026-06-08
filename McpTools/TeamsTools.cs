using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class TeamsTools
{
    [McpServerTool(Name = "teams_list"), Description("List Teams that the current user has joined")]
    public static async Task<string> ListTeams(
        [Description("Number of teams to retrieve (default: 25)")] int top = 25)
    {
        try
        {
            var result = await TeamsService.ListTeamsAsync(top);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "teams_channels_list"), Description("List channels in a Teams team")]
    public static async Task<string> ListChannels(
        [Description("Team ID")] string teamId,
        [Description("Number of channels to retrieve (default: 25)")] int top = 25)
    {
        try
        {
            var result = await TeamsService.ListChannelsAsync(teamId, top);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "teams_channel_messages"), Description("List messages in a Teams channel")]
    public static async Task<string> ListMessages(
        [Description("Team ID")] string teamId,
        [Description("Channel ID")] string channelId,
        [Description("Number of messages to retrieve (default: 20)")] int top = 20)
    {
        try
        {
            var result = await TeamsService.ListMessagesAsync(teamId, channelId, top);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "teams_channel_send"), Description("Send a new message to a Teams channel. To @-mention users, set contentType=html, include <at id=\"N\">Name</at> tags in the body, and pass comma-separated emails/AAD IDs in `mentions` where index N matches the at-tag id.")]
    public static async Task<string> SendMessage(
        [Description("Team ID")] string teamId,
        [Description("Channel ID")] string channelId,
        [Description("Message text")] string message,
        [Description("Content type: text or html (default: text). Required to be html when mentions is set.")] string contentType = "text",
        [Description("Comma-separated emails or AAD user IDs to @-mention. Body must contain <at id=\"N\">Name</at> tags (N is zero-based index into this list).")] string? mentions = null)
    {
        try
        {
            var mentionList = string.IsNullOrEmpty(mentions)
                ? null
                : mentions.Split(',').Select(m => m.Trim()).ToArray();
            var result = await TeamsService.SendMessageAsync(teamId, channelId, message, contentType, mentionList);
            return McpGraphHelper.ToJson(result);
        }
        catch (ArgumentException ex) { return McpGraphHelper.Error("invalid_argument", ex.Message); }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "teams_channel_reply"), Description("Reply to a message thread in a Teams channel. To @-mention users, set contentType=html, include <at id=\"N\">Name</at> tags in the body, and pass comma-separated emails/AAD IDs in `mentions`.")]
    public static async Task<string> Reply(
        [Description("Team ID")] string teamId,
        [Description("Channel ID")] string channelId,
        [Description("Message ID of the thread root to reply to")] string messageId,
        [Description("Reply text")] string message,
        [Description("Content type: text or html (default: text). Required to be html when mentions is set.")] string contentType = "text",
        [Description("Comma-separated emails or AAD user IDs to @-mention. Body must contain <at id=\"N\">Name</at> tags (N is zero-based index into this list).")] string? mentions = null)
    {
        try
        {
            var mentionList = string.IsNullOrEmpty(mentions)
                ? null
                : mentions.Split(',').Select(m => m.Trim()).ToArray();
            var result = await TeamsService.ReplyAsync(teamId, channelId, messageId, message, contentType, mentionList);
            return McpGraphHelper.ToJson(result);
        }
        catch (ArgumentException ex) { return McpGraphHelper.Error("invalid_argument", ex.Message); }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
