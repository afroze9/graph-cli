using Microsoft.Graph.Communications.GetPresencesByUserId;

namespace GraphCli.Services;

public static class PresenceService
{
    public static async Task<object> GetMeAsync()
    {
        var client = await GraphClientProvider.CreateAsync();
        var presence = await client.Me.Presence.GetAsync();
        return new
        {
            presence!.Availability,
            presence.Activity,
            StatusMessage = presence.StatusMessage?.Message?.Content
        };
    }

    public static async Task<object> GetAsync(string userId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var presence = await client.Communications.Presences[userId].GetAsync();
        return new
        {
            presence!.Id,
            presence.Availability,
            presence.Activity
        };
    }

    public static async Task<object> BatchAsync(string userIds)
    {
        var client = await GraphClientProvider.CreateAsync();
        var ids = userIds.Split(',').Select(id => id.Trim()).ToList();
        var presences = await client.Communications.GetPresencesByUserId
            .PostAsGetPresencesByUserIdPostResponseAsync(
                new GetPresencesByUserIdPostRequestBody { Ids = ids });
        return presences?.Value?.Select(p => new
        {
            p.Id,
            p.Availability,
            p.Activity
        }).ToList() ?? [];
    }
}
