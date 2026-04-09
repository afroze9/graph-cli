namespace GraphCli.Services;

public static class SiteService
{
    public static async Task<object> SearchAsync(string query, int top, bool refresh)
    {
        // Check cache first
        if (!refresh)
        {
            var cached = SiteCacheService.Search(query, top);
            if (cached.Count > 0)
            {
                return cached.Select(s => new
                {
                    s.Id,
                    s.Name,
                    s.DisplayName,
                    s.WebUrl,
                    Source = "cache"
                }).ToList();
            }
        }

        var client = await GraphClientProvider.CreateAsync();

        var sites = await client.Sites.GetAsync(r =>
        {
            r.QueryParameters.Search = query;
            r.QueryParameters.Top = top;
            r.QueryParameters.Select = ["id", "name", "displayName", "webUrl", "description"];
        });

        // Cache results
        if (sites?.Value != null)
        {
            SiteCacheService.UpsertMany(sites.Value
                .Where(s => s.Id != null && s.Name != null)
                .Select(s => (s.Id!, s.Name!, s.DisplayName, s.WebUrl)));
        }

        return sites?.Value?.Select(s => new
        {
            s.Id,
            s.Name,
            s.DisplayName,
            s.Description,
            s.WebUrl
        }).ToList() ?? [];
    }
}
