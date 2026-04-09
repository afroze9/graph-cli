namespace GraphCli.Services;

public static class ListService
{
    public static async Task<object> ListAsync(string site)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await PageService.ResolveSiteIdAsync(client, site);

        var lists = await client.Sites[siteId].Lists.GetAsync(r =>
        {
            r.QueryParameters.Select = ["id", "name", "displayName", "description", "webUrl",
                "createdDateTime", "lastModifiedDateTime", "list"];
        });

        return lists?.Value?.Select(l => new
        {
            l.Id,
            l.Name,
            l.DisplayName,
            l.Description,
            Template = l.ListProp?.Template?.ToString(),
            Hidden = l.ListProp?.Hidden,
            l.WebUrl,
            l.CreatedDateTime,
            l.LastModifiedDateTime
        }).ToList() ?? [];
    }

    public static async Task<object> ItemsAsync(string site, string listId, int top, string? fields, string? filter)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await PageService.ResolveSiteIdAsync(client, site);

        var items = await client.Sites[siteId].Lists[listId].Items.GetAsync(r =>
        {
            r.QueryParameters.Top = top;

            if (!string.IsNullOrEmpty(fields))
                r.QueryParameters.Expand = [$"fields(select={fields})"];
            else
                r.QueryParameters.Expand = ["fields"];

            if (!string.IsNullOrEmpty(filter))
                r.QueryParameters.Filter = filter;
        });

        return items?.Value?.Select(i => new
        {
            i.Id,
            i.WebUrl,
            i.CreatedDateTime,
            i.LastModifiedDateTime,
            CreatedBy = i.CreatedBy?.User?.DisplayName,
            LastModifiedBy = i.LastModifiedBy?.User?.DisplayName,
            Fields = i.Fields?.AdditionalData
        }).ToList() ?? [];
    }
}
