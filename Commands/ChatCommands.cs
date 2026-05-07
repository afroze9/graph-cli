using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class ChatCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var chatCommand = new Command("chat", "Chat operations");

        chatCommand.Subcommands.Add(BuildList(formatOption));
        chatCommand.Subcommands.Add(BuildSearch(formatOption));
        chatCommand.Subcommands.Add(BuildFindWith(formatOption));
        chatCommand.Subcommands.Add(BuildGet(formatOption));
        chatCommand.Subcommands.Add(BuildCreate(formatOption));
        chatCommand.Subcommands.Add(BuildMembers(formatOption));
        chatCommand.Subcommands.Add(BuildMessages(formatOption));
        chatCommand.Subcommands.Add(BuildSend(formatOption));
        chatCommand.Subcommands.Add(BuildReply(formatOption));
        chatCommand.Subcommands.Add(BuildDownloadHostedContent(formatOption));

        return chatCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 20, Description = "Number of chats to retrieve" };
        var cmd = new Command("list", "List chats") { topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var top = parseResult.GetValue(topOption);
            try
            {
                var result = await ChatService.ListAsync(top);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildSearch(Option<string> formatOption)
    {
        var queryOption = new Option<string>("--query") { Description = "Search text (case-insensitive match against chat topic)", Required = true };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 20, Description = "Max results to return" };
        var refreshOption = new Option<bool>("--refresh") { DefaultValueFactory = _ => false, Description = "Skip cache and search via API" };
        var cmd = new Command("search", "Search chats by topic") { queryOption, topOption, refreshOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var query = parseResult.GetValue(queryOption)!;
            var top = parseResult.GetValue(topOption);
            var refresh = parseResult.GetValue(refreshOption);
            try
            {
                var result = await ChatService.SearchAsync(query, top, refresh);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildFindWith(Option<string> formatOption)
    {
        var userOption = new Option<string>("--user") { Description = "User email, UPN, or AAD object ID", Required = true };
        var typeOption = new Option<string>("--type") { DefaultValueFactory = _ => "all", Description = "Chat type filter: oneOnOne, group, or all" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 20, Description = "Max results to return" };
        var cmd = new Command("find-with", "Find chats whose members include a specific user (server-side filter)") { userOption, typeOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var user = parseResult.GetValue(userOption)!;
            var type = parseResult.GetValue(typeOption) ?? "all";
            var top = parseResult.GetValue(topOption);
            try
            {
                var result = await ChatService.FindWithAsync(user, type, top);
                OutputService.Print(result, format);
            }
            catch (ArgumentException ex)
            {
                OutputService.PrintError("invalid_argument", ex.Message);
                Environment.ExitCode = 1;
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildGet(Option<string> formatOption)
    {
        var chatIdArg = new Argument<string>("chat-id") { Description = "Chat ID" };
        var cmd = new Command("get", "Get chat details") { chatIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var chatId = parseResult.GetValue(chatIdArg)!;
            try
            {
                var result = await ChatService.GetAsync(chatId);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildCreate(Option<string> formatOption)
    {
        var membersOption = new Option<string>("--members") { Description = "Comma-separated member emails", Required = true };
        var topicOption = new Option<string?>("--topic") { Description = "Chat topic (for group chats)" };
        var typeOption = new Option<string>("--type") { DefaultValueFactory = _ => "oneOnOne", Description = "Chat type: oneOnOne or group" };
        var cmd = new Command("create", "Create a new chat") { membersOption, topicOption, typeOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var members = parseResult.GetValue(membersOption)!;
            var topic = parseResult.GetValue(topicOption);
            var type = parseResult.GetValue(typeOption)!;

            var memberEmails = members.Split(',').Select(e => e.Trim());
            if (!AllowedContactsService.CheckAllAndPrompt(memberEmails, "chat"))
            {
                Environment.ExitCode = 1;
                return;
            }

            try
            {
                var result = await ChatService.CreateAsync(members, topic, type);
                OutputService.Print(result);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildMembers(Option<string> formatOption)
    {
        var chatIdArg = new Argument<string>("chat-id") { Description = "Chat ID" };
        var cmd = new Command("members", "List chat members") { chatIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var chatId = parseResult.GetValue(chatIdArg)!;
            try
            {
                var result = await ChatService.MembersAsync(chatId);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildMessages(Option<string> formatOption)
    {
        var chatIdArg = new Argument<string>("chat-id") { Description = "Chat ID" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 20, Description = "Number of messages" };
        var cmd = new Command("messages", "List chat messages") { chatIdArg, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var chatId = parseResult.GetValue(chatIdArg)!;
            var top = parseResult.GetValue(topOption);
            try
            {
                var result = await ChatService.MessagesAsync(chatId, top);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildSend(Option<string> formatOption)
    {
        var chatIdArg = new Argument<string>("chat-id") { Description = "Chat ID" };
        var messageOption = new Option<string>("--message") { Description = "Message text", Required = true };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Content type: text or html" };
        var mentionsOption = new Option<string?>("--mentions") { Description = "Comma-separated user emails or AAD IDs to @-mention (requires --content-type html; body must contain <at id=\"N\">Name</at> tags)" };
        var cmd = new Command("send", "Send a chat message") { chatIdArg, messageOption, contentTypeOption, mentionsOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var chatId = parseResult.GetValue(chatIdArg)!;
            var message = parseResult.GetValue(messageOption)!;
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";
            var mentionsCsv = parseResult.GetValue(mentionsOption);

            if (!AllowedContactsService.CheckAndPrompt(chatId, "chat"))
            {
                Environment.ExitCode = 1;
                return;
            }

            try
            {
                var mentions = string.IsNullOrEmpty(mentionsCsv)
                    ? null
                    : mentionsCsv.Split(',').Select(m => m.Trim()).ToArray();
                var result = await ChatService.SendAsync(chatId, message, contentType, mentions);
                OutputService.Print(result);
            }
            catch (ArgumentException ex)
            {
                OutputService.PrintError("invalid_argument", ex.Message);
                Environment.ExitCode = 1;
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildReply(Option<string> formatOption)
    {
        var chatIdArg = new Argument<string>("chat-id") { Description = "Chat ID" };
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID to reply to" };
        var messageOption = new Option<string>("--message") { Description = "Reply text", Required = true };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Content type: text or html" };
        var mentionsOption = new Option<string?>("--mentions") { Description = "Comma-separated user emails or AAD IDs to @-mention (requires --content-type html; body must contain <at id=\"N\">Name</at> tags)" };
        var cmd = new Command("reply", "Reply to a chat message") { chatIdArg, messageIdArg, messageOption, contentTypeOption, mentionsOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var chatId = parseResult.GetValue(chatIdArg)!;
            var messageId = parseResult.GetValue(messageIdArg)!;
            var message = parseResult.GetValue(messageOption)!;
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";
            var mentionsCsv = parseResult.GetValue(mentionsOption);

            if (!AllowedContactsService.CheckAndPrompt(chatId, "chat"))
            {
                Environment.ExitCode = 1;
                return;
            }

            try
            {
                var mentions = string.IsNullOrEmpty(mentionsCsv)
                    ? null
                    : mentionsCsv.Split(',').Select(m => m.Trim()).ToArray();
                var result = await ChatService.ReplyAsync(chatId, messageId, message, contentType, mentions);
                OutputService.Print(result);
            }
            catch (ArgumentException ex)
            {
                OutputService.PrintError("invalid_argument", ex.Message);
                Environment.ExitCode = 1;
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildDownloadHostedContent(Option<string> formatOption)
    {
        var chatIdArg = new Argument<string>("chat-id") { Description = "Chat ID" };
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID containing the hosted content" };
        var hostedContentIdArg = new Argument<string>("hosted-content-id") { Description = "Hosted content ID (from the audio/image card URL)" };
        var outOption = new Option<string>("--out") { Description = "Output file path", Required = true };
        var cmd = new Command("download-hosted-content", "Download a hosted content (image/audio) attached to a chat message") { chatIdArg, messageIdArg, hostedContentIdArg, outOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var chatId = parseResult.GetValue(chatIdArg)!;
            var messageId = parseResult.GetValue(messageIdArg)!;
            var hostedContentId = parseResult.GetValue(hostedContentIdArg)!;
            var outPath = parseResult.GetValue(outOption)!;
            try
            {
                var result = await ChatService.DownloadHostedContentAsync(chatId, messageId, hostedContentId, outPath);
                OutputService.Print(result);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }
}
