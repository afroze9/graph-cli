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
        chatCommand.Subcommands.Add(BuildGet(formatOption));
        chatCommand.Subcommands.Add(BuildCreate(formatOption));
        chatCommand.Subcommands.Add(BuildMembers(formatOption));
        chatCommand.Subcommands.Add(BuildMessages(formatOption));
        chatCommand.Subcommands.Add(BuildSend(formatOption));
        chatCommand.Subcommands.Add(BuildReply(formatOption));

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
                var client = await GraphClientProvider.CreateAsync();
                var chats = await client.Me.Chats.GetAsync(r =>
                {
                    r.QueryParameters.Top = top;
                    r.QueryParameters.Select = ["id", "topic", "chatType", "createdDateTime", "lastUpdatedDateTime"];
                    r.QueryParameters.Orderby = ["lastUpdatedDateTime desc"];
                }, ct);
                var results = chats?.Value?.Select(c => new
                {
                    c.Id,
                    c.Topic,
                    ChatType = c.ChatType?.ToString(),
                    c.CreatedDateTime,
                    c.LastUpdatedDateTime
                }).ToList();
                OutputService.Print(results, format);
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
                var client = await GraphClientProvider.CreateAsync();
                var chat = await client.Me.Chats[chatId].GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id", "topic", "chatType", "createdDateTime", "lastUpdatedDateTime", "webUrl"];
                }, ct);
                OutputService.Print(new
                {
                    chat!.Id,
                    chat.Topic,
                    ChatType = chat.ChatType?.ToString(),
                    chat.CreatedDateTime,
                    chat.LastUpdatedDateTime,
                    chat.WebUrl
                }, format);
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
                var client = await GraphClientProvider.CreateAsync();

                // Get current user's ID for the member list
                var me = await client.Me.GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id"];
                }, ct);

                var memberEmailList = members.Split(',').Select(e => e.Trim()).ToList();
                var chatMembers = new List<ConversationMember>();

                // Add current user
                chatMembers.Add(new AadUserConversationMember
                {
                    Roles = ["owner"],
                    AdditionalData = new Dictionary<string, object>
                    {
                        ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{me!.Id}')"
                    }
                });

                // Add other members
                foreach (var email in memberEmailList)
                {
                    chatMembers.Add(new AadUserConversationMember
                    {
                        Roles = ["owner"],
                        AdditionalData = new Dictionary<string, object>
                        {
                            ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{email}')"
                        }
                    });
                }

                var chat = new Chat
                {
                    ChatType = type == "group" ? ChatType.Group : ChatType.OneOnOne,
                    Topic = topic,
                    Members = chatMembers
                };

                var created = await client.Chats.PostAsync(chat, cancellationToken: ct);
                OutputService.Print(new { status = "created", id = created?.Id, chatType = type, topic });
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
                var client = await GraphClientProvider.CreateAsync();
                var members = await client.Me.Chats[chatId].Members.GetAsync(cancellationToken: ct);
                var results = members?.Value?.Select(m => new
                {
                    m.Id,
                    m.DisplayName,
                    m.Roles,
                    Email = (m as AadUserConversationMember)?.Email
                }).ToList();
                OutputService.Print(results, format);
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
                var client = await GraphClientProvider.CreateAsync();
                var messages = await client.Me.Chats[chatId].Messages.GetAsync(r =>
                {
                    r.QueryParameters.Top = top;
                    r.QueryParameters.Orderby = ["createdDateTime desc"];
                }, ct);
                var results = messages?.Value?.Select(m => new
                {
                    m.Id,
                    From = m.From?.User?.DisplayName ?? m.From?.Application?.DisplayName,
                    BodyType = m.Body?.ContentType?.ToString(),
                    Body = m.Body?.Content,
                    m.CreatedDateTime,
                    MessageType = m.MessageType?.ToString()
                }).ToList();
                OutputService.Print(results, format);
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
        var cmd = new Command("send", "Send a chat message") { chatIdArg, messageOption, contentTypeOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var chatId = parseResult.GetValue(chatIdArg)!;
            var message = parseResult.GetValue(messageOption)!;
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";

            try
            {
                var client = await GraphClientProvider.CreateAsync();

                var chatMembers = await client.Me.Chats[chatId].Members.GetAsync(cancellationToken: ct);
                var me = await client.Me.GetAsync(r => r.QueryParameters.Select = ["mail"], ct);
                var otherMembers = chatMembers?.Value?
                    .OfType<AadUserConversationMember>()
                    .Where(m => !string.Equals(m.Email, me?.Mail, StringComparison.OrdinalIgnoreCase))
                    .Select(m => m.Email ?? m.DisplayName ?? "unknown")
                    .Where(e => !string.IsNullOrEmpty(e))
                    .ToList() ?? [];

                if (!AllowedContactsService.CheckAllAndPrompt(otherMembers, "chat"))
                {
                    Environment.ExitCode = 1;
                    return;
                }

                var chatMessage = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        ContentType = contentType == "html" ? BodyType.Html : BodyType.Text,
                        Content = message
                    }
                };
                var sent = await client.Me.Chats[chatId].Messages.PostAsync(chatMessage, cancellationToken: ct);
                OutputService.Print(new { status = "sent", id = sent?.Id, chatId });
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
        var cmd = new Command("reply", "Reply to a chat message") { chatIdArg, messageIdArg, messageOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var chatId = parseResult.GetValue(chatIdArg)!;
            var messageId = parseResult.GetValue(messageIdArg)!;
            var message = parseResult.GetValue(messageOption)!;

            try
            {
                var client = await GraphClientProvider.CreateAsync();

                var chatMembers = await client.Me.Chats[chatId].Members.GetAsync(cancellationToken: ct);
                var me = await client.Me.GetAsync(r => r.QueryParameters.Select = ["mail"], ct);
                var otherMembers = chatMembers?.Value?
                    .OfType<AadUserConversationMember>()
                    .Where(m => !string.Equals(m.Email, me?.Mail, StringComparison.OrdinalIgnoreCase))
                    .Select(m => m.Email ?? m.DisplayName ?? "unknown")
                    .Where(e => !string.IsNullOrEmpty(e))
                    .ToList() ?? [];

                if (!AllowedContactsService.CheckAllAndPrompt(otherMembers, "chat"))
                {
                    Environment.ExitCode = 1;
                    return;
                }

                var reply = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = message
                    }
                };
                var sent = await client.Me.Chats[chatId].Messages[messageId].Replies.PostAsync(reply, cancellationToken: ct);
                OutputService.Print(new { status = "replied", id = sent?.Id, chatId, messageId });
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
