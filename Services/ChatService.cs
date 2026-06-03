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

    public static async Task<object> SearchAsync(string query, int top, bool refresh, int maxDepth)
    {
        // Fast path: search cache by topic OR member display name OR email.
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
                    c.Members,
                    c.LastUpdatedDateTime,
                    c.LastUsed,
                    Source = "cache"
                }).ToList();
            }
        }

        // Refresh path: fetch up to maxDepth most-recently-active chats with
        // members expanded, merge into the cache, then re-run the cache search.
        await RefreshAsync(maxDepth);

        return ChatCacheService.Search(query, top).Select(c => new
        {
            c.Id,
            Topic = c.Name,
            c.ChatType,
            c.Members,
            c.LastUpdatedDateTime,
            c.LastUsed,
            Source = refresh ? "api (refresh)" : "api (cache miss)"
        }).ToList();
    }

    /// <summary>
    /// Paginated pull of Me/Chats with $expand=members, ordered by most recent
    /// activity. Merges into the on-disk chat cache. Stops once <paramref name="maxDepth"/>
    /// chats have been fetched or pagination ends.
    /// </summary>
    /// <remarks>
    /// Graph caps $top at 50 per page and caps expanded members at 25 per chat
    /// regardless of $top. The 25-member cap is acceptable for 1:1 and small
    /// group chats; large groups will only have their first 25 members cached.
    /// </remarks>
    private static async Task RefreshAsync(int maxDepth)
    {
        var client = await GraphClientProvider.CreateAsync();
        var fetched = 0;

        var page = await client.Me.Chats.GetAsync(r =>
        {
            r.QueryParameters.Top = Math.Min(50, maxDepth);
            r.QueryParameters.Expand = ["members"];
            r.QueryParameters.Orderby = ["lastMessagePreview/createdDateTime desc"];
            r.QueryParameters.Select = ["id", "topic", "chatType", "createdDateTime", "lastUpdatedDateTime"];
        });

        while (page?.Value != null)
        {
            var entries = page.Value
                .Where(c => c.Id != null)
                .Select(c => new ChatCacheEntry(
                    c.Id!,
                    c.Topic,
                    c.ChatType?.ToString(),
                    c.Members?.Select(m => new CachedMember
                    {
                        DisplayName = m.DisplayName,
                        Email = (m as AadUserConversationMember)?.Email,
                        UserId = (m as AadUserConversationMember)?.UserId
                    }).ToList(),
                    c.LastUpdatedDateTime))
                .ToList();

            ChatCacheService.UpsertMany(entries);
            fetched += entries.Count;

            if (fetched >= maxDepth || string.IsNullOrEmpty(page.OdataNextLink))
                break;

            page = await client.Me.Chats.WithUrl(page.OdataNextLink).GetAsync();
        }
    }

    public static async Task<object> FindWithAsync(string user, string? type, int top)
    {
        var client = await GraphClientProvider.CreateAsync();

        Microsoft.Graph.Models.User? resolved;
        try
        {
            resolved = await client.Users[user].GetAsync(r =>
            {
                r.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName"];
            });
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            throw new ArgumentException($"could not resolve user '{user}': {ex.Error?.Message ?? ex.Message}");
        }

        if (resolved?.Id == null)
            throw new ArgumentException($"user '{user}' not found");

        var memberFilter = $"members/any(o:o/microsoft.graph.aadUserConversationMember/userId eq '{resolved.Id}')";
        var filter = type switch
        {
            "oneOnOne" => $"chatType eq 'oneOnOne' and {memberFilter}",
            "group" => $"chatType eq 'group' and {memberFilter}",
            _ => memberFilter
        };

        var chats = await client.Me.Chats.GetAsync(r =>
        {
            r.QueryParameters.Top = top;
            r.QueryParameters.Filter = filter;
            r.QueryParameters.Expand = ["members"];
            r.QueryParameters.Select = ["id", "topic", "chatType", "createdDateTime", "lastUpdatedDateTime"];
        });

        if (chats?.Value != null)
        {
            ChatCacheService.UpsertMany(chats.Value
                .Where(c => c.Id != null)
                .Select(c => (c.Id!, c.Topic, c.ChatType?.ToString())));
        }

        return new
        {
            user = new { resolved.Id, resolved.DisplayName, Email = resolved.Mail ?? resolved.UserPrincipalName },
            chats = chats?.Value?.Select(c => new
            {
                c.Id,
                c.Topic,
                ChatType = c.ChatType?.ToString(),
                c.CreatedDateTime,
                c.LastUpdatedDateTime,
                Members = c.Members?.Select(m => new
                {
                    m.DisplayName,
                    Email = (m as AadUserConversationMember)?.Email
                }).ToList()
            }).ToList()
        };
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
            MessageType = m.MessageType?.ToString(),
            Attachments = m.Attachments?.Select(a => new
            {
                a.Id,
                a.ContentType,
                a.Name,
                a.ContentUrl,
                a.Content
            }).ToList()
        }).ToList()!;
    }

    public static async Task<object> DownloadHostedContentAsync(string chatId, string messageId, string hostedContentId, string outPath)
    {
        var client = await GraphClientProvider.CreateAsync();
        var stream = await client.Me.Chats[chatId].Messages[messageId].HostedContents[hostedContentId].Content.GetAsync();
        if (stream == null)
        {
            return new { status = "error", message = "no content" };
        }
        await using var fs = File.Create(outPath);
        await stream.CopyToAsync(fs);
        return new { status = "downloaded", file = outPath, size = new FileInfo(outPath).Length };
    }

    public static async Task<object> SendImageAsync(string chatId, string imagePath, string? caption = null)
    {
        if (!File.Exists(imagePath))
            throw new ArgumentException($"Image file not found: {imagePath}");

        var ext = Path.GetExtension(imagePath).TrimStart('.').ToLowerInvariant();
        var mimeType = ext switch
        {
            "png"           => "image/png",
            "jpg" or "jpeg" => "image/jpeg",
            "gif"           => "image/gif",
            "webp"          => "image/webp",
            _               => throw new ArgumentException(
                $"Unsupported image type '.{ext}'. Supported types: png, jpg, jpeg, gif, webp.")
        };

        var imageBytes = await File.ReadAllBytesAsync(imagePath);

        // Teams inline hosted content is capped at ~4 MB; fail early with a clear
        // message rather than letting Graph reject it with an opaque error.
        const int maxInlineBytes = 4 * 1024 * 1024;
        if (imageBytes.Length > maxInlineBytes)
            throw new ArgumentException(
                $"Image is {imageBytes.Length / (1024 * 1024.0):0.0} MB, which exceeds the ~4 MB inline limit. " +
                "Share it as a file instead (e.g. 'files share').");

        var client = await GraphClientProvider.CreateAsync();

        // Build body: optional caption text above the inline image
        var bodyHtml = string.IsNullOrWhiteSpace(caption)
            ? "<img src=\"../hostedContents/1/$value\" style=\"vertical-align:bottom;max-width:800px;\" />"
            : $"<p>{System.Net.WebUtility.HtmlEncode(caption)}</p><img src=\"../hostedContents/1/$value\" style=\"vertical-align:bottom;max-width:800px;\" />";

        var chatMessage = new ChatMessage
        {
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = bodyHtml
            },
            HostedContents = new List<ChatMessageHostedContent>
            {
                new ChatMessageHostedContent
                {
                    ContentBytes = imageBytes,
                    ContentType = mimeType,
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.temporaryId", "1" }
                    }
                }
            }
        };

        var sent = await client.Me.Chats[chatId].Messages.PostAsync(chatMessage);
        ChatCacheService.Upsert(chatId, null, null);
        return new { status = "sent", id = sent?.Id, chatId, mimeType };
    }

    public static async Task<object> SendAsync(string chatId, string message, string contentType, string[]? mentions = null)
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

        if (mentions is { Length: > 0 })
            chatMessage.Mentions = await BuildMentionsAsync(client, message, contentType, mentions);

        var sent = await client.Me.Chats[chatId].Messages.PostAsync(chatMessage);
        ChatCacheService.Upsert(chatId, null, null);
        return new { status = "sent", id = sent?.Id, chatId };
    }

    public static async Task<object> ReplyAsync(string chatId, string messageId, string message, string contentType, string[]? mentions = null)
    {
        var client = await GraphClientProvider.CreateAsync();
        var reply = new ChatMessage
        {
            Body = new ItemBody
            {
                ContentType = contentType == "html" ? BodyType.Html : BodyType.Text,
                Content = message
            }
        };

        if (mentions is { Length: > 0 })
            reply.Mentions = await BuildMentionsAsync(client, message, contentType, mentions);

        var sent = await client.Me.Chats[chatId].Messages[messageId].Replies.PostAsync(reply);
        ChatCacheService.Upsert(chatId, null, null);
        return new { status = "replied", id = sent?.Id, chatId, messageId };
    }

    // Build a ChatMessage.Mentions collection from a list of user emails or AAD IDs.
    // Body must be HTML and must contain `<at id="N">Name</at>` tags for each mention
    // (zero-based). Resolves each identifier to an AAD user and sets userIdentityType=aadUser
    // so Teams fires a notification rather than rendering the @-tag as plain text.
    private static async Task<List<ChatMessageMention>> BuildMentionsAsync(
        GraphServiceClient client, string body, string contentType, string[] mentions)
    {
        if (!string.Equals(contentType, "html", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                "--mentions requires --content-type html. Reference each mention in the body with <at id=\"N\">Name</at> where N is the zero-based index matching --mentions.");

        var result = new List<ChatMessageMention>();
        for (int i = 0; i < mentions.Length; i++)
        {
            var identifier = mentions[i].Trim();

            // Graph requires the `<at id="N">Display</at>` tag in the body, and the inner
            // text must match the mention's MentionText. Extract the inner text from the
            // body so callers can write any display text they like (e.g. "Ali") without
            // having to match the user's full AAD displayName.
            var atMatch = System.Text.RegularExpressions.Regex.Match(
                body,
                $"<at id=[\"']{i}[\"'][^>]*>(.*?)</at>",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
            if (!atMatch.Success)
                throw new ArgumentException(
                    $"--mentions: body is missing <at id=\"{i}\">...</at> tag for mention #{i} ({identifier}). Add it to the --message HTML so Teams can render the @-mention.");
            var mentionText = atMatch.Groups[1].Value.Trim();

            Microsoft.Graph.Models.User? user;
            try
            {
                user = await client.Users[identifier].GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id", "displayName"];
                });
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                throw new ArgumentException(
                    $"--mentions: could not resolve user '{identifier}': {ex.Error?.Message ?? ex.Message}");
            }

            if (user?.Id == null)
                throw new ArgumentException($"--mentions: user '{identifier}' not found");

            var identity = new Identity
            {
                Id = user.Id,
                DisplayName = user.DisplayName
            };
            // userIdentityType is not a typed property on Identity; Graph expects it
            // as an extension field in the JSON — populate via AdditionalData.
            identity.AdditionalData["userIdentityType"] = "aadUser";

            result.Add(new ChatMessageMention
            {
                Id = i,
                MentionText = mentionText,
                Mentioned = new ChatMessageMentionedIdentitySet
                {
                    User = identity
                }
            });
        }
        return result;
    }
}
