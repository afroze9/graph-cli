using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace GraphCli.Services;

public static class TeamsService
{
    // -------------------------------------------------------------------------
    // Teams & Channels — discovery
    // -------------------------------------------------------------------------

    public static async Task<object> ListTeamsAsync(int top)
    {
        var client = await GraphClientProvider.CreateAsync();
        var teams = await client.Me.JoinedTeams.GetAsync(r =>
        {
            r.QueryParameters.Top = top;
            r.QueryParameters.Select = ["id", "displayName", "description", "isArchived"];
        });

        return teams?.Value?.Select(t => new
        {
            t.Id,
            t.DisplayName,
            t.Description,
            t.IsArchived
        }).ToList() ?? [];
    }

    public static async Task<object> ListChannelsAsync(string teamId, int top)
    {
        var client = await GraphClientProvider.CreateAsync();
        var channels = await client.Teams[teamId].Channels.GetAsync(r =>
        {
            r.QueryParameters.Top = top;
            r.QueryParameters.Select = ["id", "displayName", "description", "membershipType", "webUrl"];
        });

        return channels?.Value?.Select(c => new
        {
            c.Id,
            c.DisplayName,
            c.Description,
            MembershipType = c.MembershipType?.ToString(),
            c.WebUrl
        }).ToList() ?? [];
    }

    // -------------------------------------------------------------------------
    // Channel Messages
    // -------------------------------------------------------------------------

    public static async Task<object> ListMessagesAsync(string teamId, string channelId, int top)
    {
        var client = await GraphClientProvider.CreateAsync();
        var messages = await client.Teams[teamId].Channels[channelId].Messages.GetAsync(r =>
        {
            r.QueryParameters.Top = top;
            r.QueryParameters.Select = ["id", "subject", "body", "from", "createdDateTime", "lastModifiedDateTime", "replyToId", "webUrl"];
        });

        return messages?.Value?.Select(m => new
        {
            m.Id,
            m.Subject,
            Body = m.Body?.Content,
            ContentType = m.Body?.ContentType?.ToString(),
            From = m.From?.User?.DisplayName ?? m.From?.Application?.DisplayName,
            m.CreatedDateTime,
            m.LastModifiedDateTime,
            m.ReplyToId,
            m.WebUrl
        }).ToList() ?? [];
    }

    // -------------------------------------------------------------------------
    // Send a new message to a channel
    // -------------------------------------------------------------------------

    public static async Task<object> SendMessageAsync(
        string teamId, string channelId, string message, string contentType, string[]? mentions)
    {
        var client = await GraphClientProvider.CreateAsync();
        var msg = BuildMessage(message, contentType);

        if (mentions is { Length: > 0 })
            msg.Mentions = await BuildMentionsAsync(client, message, contentType, mentions);

        var sent = await client.Teams[teamId].Channels[channelId].Messages.PostAsync(msg);
        return new { status = "sent", id = sent?.Id, teamId, channelId };
    }

    // -------------------------------------------------------------------------
    // Reply to an existing channel message (thread reply)
    // -------------------------------------------------------------------------

    public static async Task<object> ReplyAsync(
        string teamId, string channelId, string messageId, string message, string contentType, string[]? mentions)
    {
        var client = await GraphClientProvider.CreateAsync();
        var reply = BuildMessage(message, contentType);

        if (mentions is { Length: > 0 })
            reply.Mentions = await BuildMentionsAsync(client, message, contentType, mentions);

        var sent = await client.Teams[teamId].Channels[channelId].Messages[messageId].Replies.PostAsync(reply);
        return new { status = "replied", id = sent?.Id, teamId, channelId, messageId };
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static ChatMessage BuildMessage(string content, string contentType) => new()
    {
        Body = new ItemBody
        {
            ContentType = contentType == "html" ? BodyType.Html : BodyType.Text,
            Content = content
        }
    };

    /// <summary>
    /// Resolves a list of user emails / AAD IDs into a ChatMessage.Mentions
    /// collection, matching the same pattern used by ChatService.
    /// </summary>
    private static async Task<List<ChatMessageMention>> BuildMentionsAsync(
        GraphServiceClient client, string body, string contentType, string[] mentions)
    {
        if (contentType != "html")
            throw new ArgumentException("Mentions require contentType=html and <at id=\"N\">Name</at> tags in the body.");

        var result = new List<ChatMessageMention>();
        for (var i = 0; i < mentions.Length; i++)
        {
            var identifier = mentions[i];
            Microsoft.Graph.Models.User? user;
            try
            {
                user = await client.Users[identifier].GetAsync(r =>
                    r.QueryParameters.Select = ["id", "displayName", "mail"]);
            }
            catch
            {
                throw new ArgumentException($"Could not resolve user '{identifier}' for mention.");
            }

            result.Add(new ChatMessageMention
            {
                Id = i,
                MentionText = user?.DisplayName ?? identifier,
                Mentioned = new ChatMessageMentionedIdentitySet
                {
                    User = new Identity
                    {
                        Id = user?.Id,
                        DisplayName = user?.DisplayName,
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "userIdentityType", "aadUser" }
                        }
                    }
                }
            });
        }

        return result;
    }
}
