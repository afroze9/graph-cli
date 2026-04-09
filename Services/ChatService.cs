using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace GraphCli.Services;

public static class ChatService
{
    public static async Task<object> ListAsync(int top)
    {
        var client = await GraphClientProvider.CreateAsync();
        var chats = await client.Me.Chats.GetAsync(r =>
        {
            r.QueryParameters.Top = top;
            r.QueryParameters.Select = ["id", "topic", "chatType", "createdDateTime", "lastUpdatedDateTime"];
        });
        var results = chats?.Value?.Select(c => new
        {
            c.Id,
            c.Topic,
            ChatType = c.ChatType?.ToString(),
            c.CreatedDateTime,
            c.LastUpdatedDateTime
        }).ToList();

        if (chats?.Value != null)
        {
            ChatCacheService.UpsertMany(chats.Value
                .Where(c => c.Id != null)
                .Select(c => (c.Id!, c.Topic, c.ChatType?.ToString())));
        }

        return results!;
    }

    public static async Task<object> SearchAsync(string query, int top, bool refresh)
    {
        if (!refresh)
        {
            var cached = ChatCacheService.Search(query, top);
            if (cached.Count > 0)
            {
                return cached.Select(c => new
                {
                    c.Id,
                    Topic = c.Name,
                    c.ChatType,
                    c.LastUsed,
                    Source = "cache"
                }).ToList();
            }
        }

        var client = await GraphClientProvider.CreateAsync();
        var matched = new List<Chat>();

        var page = await client.Me.Chats.GetAsync(r =>
        {
            r.QueryParameters.Top = 50;
            r.QueryParameters.Select = ["id", "topic", "chatType", "createdDateTime", "lastUpdatedDateTime"];
        });

        while (page?.Value != null)
        {
            ChatCacheService.UpsertMany(page.Value
                .Where(c => c.Id != null)
                .Select(c => (c.Id!, c.Topic, c.ChatType?.ToString())));

            foreach (var c in page.Value)
            {
                if (c.Topic != null && c.Topic.Contains(query, StringComparison.OrdinalIgnoreCase))
                {
                    matched.Add(c);
                    if (matched.Count >= top) break;
                }
            }
            if (matched.Count >= top || string.IsNullOrEmpty(page.OdataNextLink)) break;
            page = await client.Me.Chats.WithUrl(page.OdataNextLink).GetAsync();
        }

        return matched.Select(c => new
        {
            c.Id,
            c.Topic,
            ChatType = c.ChatType?.ToString(),
            c.CreatedDateTime,
            c.LastUpdatedDateTime
        }).ToList();
    }

    public static async Task<object> GetAsync(string chatId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var chat = await client.Me.Chats[chatId].GetAsync(r =>
        {
            r.QueryParameters.Select = ["id", "topic", "chatType", "createdDateTime", "lastUpdatedDateTime", "webUrl"];
        });
        ChatCacheService.Upsert(chat!.Id!, chat.Topic, chat.ChatType?.ToString());
        return new
        {
            chat.Id,
            chat.Topic,
            ChatType = chat.ChatType?.ToString(),
            chat.CreatedDateTime,
            chat.LastUpdatedDateTime,
            chat.WebUrl
        };
    }

    public static async Task<object> CreateAsync(string members, string? topic, string type)
    {
        var client = await GraphClientProvider.CreateAsync();

        var me = await client.Me.GetAsync(r =>
        {
            r.QueryParameters.Select = ["id"];
        });

        var memberEmailList = members.Split(',').Select(e => e.Trim()).ToList();
        var chatMembers = new List<ConversationMember>();

        chatMembers.Add(new AadUserConversationMember
        {
            Roles = ["owner"],
            AdditionalData = new Dictionary<string, object>
            {
                ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{me!.Id}')"
            }
        });

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

        var created = await client.Chats.PostAsync(chat);
        if (created?.Id != null)
            ChatCacheService.Upsert(created.Id, topic, type);

        return new { status = "created", id = created?.Id, chatType = type, topic };
    }

    public static async Task<object> MembersAsync(string chatId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var members = await client.Me.Chats[chatId].Members.GetAsync();
        return members?.Value?.Select(m => new
        {
            m.Id,
            m.DisplayName,
            m.Roles,
            Email = (m as AadUserConversationMember)?.Email
        }).ToList()!;
    }

    public static async Task<object> MessagesAsync(string chatId, int top)
    {
        var client = await GraphClientProvider.CreateAsync();
        var messages = await client.Me.Chats[chatId].Messages.GetAsync(r =>
        {
            r.QueryParameters.Top = top;
            r.QueryParameters.Orderby = ["createdDateTime desc"];
        });
        return messages?.Value?.Select(m => new
        {
            m.Id,
            From = m.From?.User?.DisplayName ?? m.From?.Application?.DisplayName,
            BodyType = m.Body?.ContentType?.ToString(),
            Body = m.Body?.Content,
            m.CreatedDateTime,
            MessageType = m.MessageType?.ToString()
        }).ToList()!;
    }

    public static async Task<object> SendAsync(string chatId, string message, string contentType)
    {
        var client = await GraphClientProvider.CreateAsync();
        var chatMessage = new ChatMessage
        {
            Body = new ItemBody
            {
                ContentType = contentType == "html" ? BodyType.Html : BodyType.Text,
                Content = message
            }
        };
        var sent = await client.Me.Chats[chatId].Messages.PostAsync(chatMessage);
        ChatCacheService.Upsert(chatId, null, null);
        return new { status = "sent", id = sent?.Id, chatId };
    }

    public static async Task<object> ReplyAsync(string chatId, string messageId, string message)
    {
        var client = await GraphClientProvider.CreateAsync();
        var reply = new ChatMessage
        {
            Body = new ItemBody
            {
                ContentType = BodyType.Text,
                Content = message
            }
        };
        var sent = await client.Me.Chats[chatId].Messages[messageId].Replies.PostAsync(reply);
        ChatCacheService.Upsert(chatId, null, null);
        return new { status = "replied", id = sent?.Id, chatId, messageId };
    }
}
