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
            }).ToList(),
            // Emoji reactions on this message (e.g. a 👍 thumbs-up). reactionType is the
            // Unicode emoji (or a backward-compatible name such as "like"/"heart"); the
            // reaction's own displayName is a friendly label ("Yes", "Heart"). Graph
            // populates the reacting user's id here but leaves userDisplayName null on the
            // list-messages endpoint, so userId is the reliable identifier of who reacted
            // (resolve it to a name via user_get if needed).
            Reactions = m.Reactions?.Select(r => new
            {
                r.ReactionType,
                r.DisplayName,
                r.CreatedDateTime,
                UserId = r.User?.User?.Id,
                UserDisplayName = r.User?.User?.DisplayName
            }).ToList()
        }).ToList()!;
    }

    /// <summary>
    /// Fetch every new chat message across all Teams chats since a given point in time.
    /// Enumerates chats ordered by most-recent activity and short-circuits once a chat's
    /// last message predates the cutoff, then pages each qualifying chat's messages
    /// (newest-first) until it crosses the cutoff. Results are returned oldest-first so
    /// they can be streamed straight into a DB / task-extraction pipeline.
    /// </summary>
    /// <param name="since">
    /// Cutoff: ISO 8601 (e.g. 2026-07-15T13:00), a bare time ("1pm" = today), "today",
    /// "yesterday", or a relative offset ("-3h", "30m", "-2d", "-1w"). Ignored when
    /// <paramref name="useCache"/> is true.
    /// </param>
    /// <param name="useCache">Use the stored watermark from the previous run instead of <paramref name="since"/>.</param>
    /// <param name="maxChats">Cap on how many recently-active chats to scan.</param>
    /// <param name="includeSystem">Include system event messages (joins/leaves/renames); default only real messages.</param>
    /// <param name="excludeOwn">Drop messages authored by the signed-in user.</param>
    /// <param name="saveWatermark">Persist the newest message timestamp as the new watermark for --continue.</param>
    public static async Task<object> SinceAsync(
        string? since,
        bool useCache,
        int maxChats,
        bool includeSystem,
        bool excludeOwn,
        bool saveWatermark)
    {
        DateTimeOffset sinceUtc;
        if (useCache || string.Equals(since?.Trim(), "last", StringComparison.OrdinalIgnoreCase))
        {
            var wm = ChatCacheService.GetWatermark()
                ?? throw new ArgumentException(
                    "no cached watermark found. Run once with an explicit --since first (e.g. --since 1pm or --since -3h).");
            sinceUtc = wm;
        }
        else
        {
            if (string.IsNullOrWhiteSpace(since))
                throw new ArgumentException(
                    "--since is required. Use ISO 8601 (2026-07-15T13:00), a time ('1pm'), 'today', or a relative offset ('-3h', '-2d'). Or pass --continue to resume from the last run.");
            sinceUtc = ParseSince(since);
        }

        var client = await GraphClientProvider.CreateAsync();

        string? myId = null;
        if (excludeOwn)
        {
            var me = await client.Me.GetAsync(r => r.QueryParameters.Select = ["id"]);
            myId = me?.Id;
        }

        // Enumerate chats ordered by most-recent activity. Because the list is sorted
        // desc by the last message time, the first chat whose (non-null) last-message
        // preview predates the cutoff means every chat after it does too — stop there.
        var chatsToScan = new List<Chat>();
        var scanned = 0;
        var stop = false;

        var page = await client.Me.Chats.GetAsync(r =>
        {
            r.QueryParameters.Top = Math.Min(50, maxChats);
            r.QueryParameters.Expand = ["members", "lastMessagePreview"];
            r.QueryParameters.Orderby = ["lastMessagePreview/createdDateTime desc"];
            r.QueryParameters.Select = ["id", "topic", "chatType", "lastUpdatedDateTime"];
        });

        while (page?.Value != null && !stop)
        {
            ChatCacheService.UpsertMany(page.Value
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
                    c.LastUpdatedDateTime)));

            foreach (var chat in page.Value)
            {
                var lastMsgAt = chat.LastMessagePreview?.CreatedDateTime;
                if (lastMsgAt.HasValue && lastMsgAt.Value <= sinceUtc)
                {
                    // Sorted desc: nothing newer remains beyond this point.
                    stop = true;
                    break;
                }

                chatsToScan.Add(chat);
                scanned++;
                if (scanned >= maxChats) { stop = true; break; }
            }

            if (stop || string.IsNullOrEmpty(page.OdataNextLink))
                break;

            page = await client.Me.Chats.WithUrl(page.OdataNextLink).GetAsync();
        }

        var rows = new List<ChatSinceMessage>();
        DateTimeOffset? newWatermark = null;

        foreach (var chat in chatsToScan)
        {
            var label = ChatLabel(chat, myId);
            var chatType = chat.ChatType?.ToString();

            var msgPage = await client.Me.Chats[chat.Id].Messages.GetAsync(r =>
            {
                r.QueryParameters.Top = 50;
                r.QueryParameters.Orderby = ["createdDateTime desc"];
            });

            var chatDone = false;
            while (msgPage?.Value != null && !chatDone)
            {
                foreach (var msg in msgPage.Value)
                {
                    var created = msg.CreatedDateTime;
                    if (created.HasValue && created.Value <= sinceUtc)
                    {
                        chatDone = true;
                        break;
                    }

                    if (msg.DeletedDateTime != null)
                        continue;

                    var messageType = msg.MessageType?.ToString() ?? "message";
                    if (!includeSystem && !string.Equals(messageType, "message", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var fromId = msg.From?.User?.Id;
                    if (excludeOwn && myId != null && string.Equals(fromId, myId, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (created.HasValue && (newWatermark == null || created.Value > newWatermark.Value))
                        newWatermark = created.Value;

                    rows.Add(new ChatSinceMessage
                    {
                        ChatId = chat.Id,
                        ChatTopic = label,
                        ChatType = chatType,
                        MessageId = msg.Id,
                        From = msg.From?.User?.DisplayName ?? msg.From?.Application?.DisplayName,
                        FromUserId = fromId,
                        CreatedDateTime = created,
                        LastModifiedDateTime = msg.LastModifiedDateTime,
                        MessageType = messageType,
                        BodyType = msg.Body?.ContentType?.ToString(),
                        Body = msg.Body?.Content,
                        Text = StripHtml(msg.Body?.Content),
                        WebUrl = msg.WebUrl,
                        Attachments = msg.Attachments?.Select(a => new ChatSinceAttachment
                        {
                            Id = a.Id,
                            ContentType = a.ContentType,
                            Name = a.Name,
                            ContentUrl = a.ContentUrl
                        }).ToList()
                    });
                }

                if (chatDone || string.IsNullOrEmpty(msgPage.OdataNextLink))
                    break;

                msgPage = await client.Me.Chats[chat.Id].Messages.WithUrl(msgPage.OdataNextLink).GetAsync();
            }
        }

        // Oldest-first: natural order for appending to a DB / running task extraction.
        var ordered = rows
            .OrderBy(r => r.CreatedDateTime ?? DateTimeOffset.MinValue)
            .ToList();

        if (saveWatermark && newWatermark.HasValue)
            ChatCacheService.SaveWatermark(newWatermark.Value);

        return new
        {
            since = sinceUtc,
            chatsScanned = chatsToScan.Count,
            count = ordered.Count,
            watermark = newWatermark ?? ChatCacheService.GetWatermark(),
            messages = ordered
        };
    }

    // Human-friendly label for a chat: the topic if set, otherwise the other
    // participants' display names (excludes the signed-in user when known).
    private static string ChatLabel(Chat chat, string? myId)
    {
        if (!string.IsNullOrWhiteSpace(chat.Topic))
            return chat.Topic!;

        var others = chat.Members?
            .Where(m => myId == null || !string.Equals((m as AadUserConversationMember)?.UserId, myId, StringComparison.OrdinalIgnoreCase))
            .Select(m => m.DisplayName)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .ToList();

        if (others is { Count: > 0 })
            return string.Join(", ", others);

        return chat.ChatType?.ToString() ?? "chat";
    }

    // Cheap HTML-to-text for message bodies: drop tags, collapse whitespace, decode entities.
    private static string? StripHtml(string? html)
    {
        if (string.IsNullOrEmpty(html))
            return html;

        var noTags = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ");
        var decoded = System.Net.WebUtility.HtmlDecode(noTags);
        return System.Text.RegularExpressions.Regex.Replace(decoded, @"\s+", " ").Trim();
    }

    private static DateTimeOffset ParseSince(string input)
    {
        input = input.Trim();
        var now = DateTimeOffset.Now;

        if (string.Equals(input, "today", StringComparison.OrdinalIgnoreCase))
            return new DateTimeOffset(DateTime.Today, now.Offset).ToUniversalTime();
        if (string.Equals(input, "yesterday", StringComparison.OrdinalIgnoreCase))
            return new DateTimeOffset(DateTime.Today.AddDays(-1), now.Offset).ToUniversalTime();

        // Relative offset: -3h, 3h, -30m, -2d, -1w (m=minutes, h=hours, d=days, w=weeks).
        var rel = System.Text.RegularExpressions.Regex.Match(
            input, @"^-?(\d+)\s*([mhdw])$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (rel.Success)
        {
            var n = int.Parse(rel.Groups[1].Value);
            var span = rel.Groups[2].Value.ToLowerInvariant() switch
            {
                "m" => TimeSpan.FromMinutes(n),
                "h" => TimeSpan.FromHours(n),
                "d" => TimeSpan.FromDays(n),
                "w" => TimeSpan.FromDays(7 * n),
                _ => TimeSpan.Zero
            };
            return now.Subtract(span).ToUniversalTime();
        }

        // ISO 8601 or a bare time like "1pm" / "13:00" (assumed today, local tz).
        if (DateTimeOffset.TryParse(input, System.Globalization.CultureInfo.CurrentCulture,
                System.Globalization.DateTimeStyles.AssumeLocal, out var dto))
            return dto.ToUniversalTime();

        throw new ArgumentException(
            $"could not parse --since value '{input}'. Use ISO 8601 (2026-07-15T13:00), a time ('1pm'), 'today'/'yesterday', or a relative offset ('-3h', '-2d', '-1w').");
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

        // The Graph API does not expose a /replies endpoint for chat messages
        // (1:1 or group). The correct approach is to POST a new message to the
        // chat with a messageReference attachment, which Teams renders as a
        // native quoted reply indistinguishable from the built-in reply UX.
        var bodyContent = contentType == "html"
            ? $"<attachment id=\"{messageId}\"></attachment>{message}"
            : $"<attachment id=\"{messageId}\"></attachment><p>{System.Net.WebUtility.HtmlEncode(message)}</p>";

        var reply = new ChatMessage
        {
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = bodyContent
            },
            Attachments =
            [
                new ChatMessageAttachment
                {
                    Id = messageId,
                    ContentType = "messageReference",
                    Content = $"{{\"messageId\":\"{messageId}\"}}"
                }
            ]
        };

        if (mentions is { Length: > 0 })
            reply.Mentions = await BuildMentionsAsync(client, message, contentType, mentions);

        var sent = await client.Me.Chats[chatId].Messages.PostAsync(reply);
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

public class ChatSinceMessage
{
    public string? ChatId { get; set; }
    public string? ChatTopic { get; set; }
    public string? ChatType { get; set; }
    public string? MessageId { get; set; }
    public string? From { get; set; }
    public string? FromUserId { get; set; }
    public DateTimeOffset? CreatedDateTime { get; set; }
    public DateTimeOffset? LastModifiedDateTime { get; set; }
    public string? MessageType { get; set; }
    public string? BodyType { get; set; }
    public string? Body { get; set; }
    public string? Text { get; set; }
    public string? WebUrl { get; set; }
    public List<ChatSinceAttachment>? Attachments { get; set; }
}

public class ChatSinceAttachment
{
    public string? Id { get; set; }
    public string? ContentType { get; set; }
    public string? Name { get; set; }
    public string? ContentUrl { get; set; }
}
