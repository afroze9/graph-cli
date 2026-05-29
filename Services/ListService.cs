using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Serialization;

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

    public static async Task<object> ItemsAsync(string site, string listId, int top, string? fields, string? filter, string? expandLookups)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await PageService.ResolveSiteIdAsync(client, site);

        var selectClause = BuildFieldsSelect(fields, expandLookups);

        var items = await client.Sites[siteId].Lists[listId].Items.GetAsync(r =>
        {
            r.QueryParameters.Top = top;

            r.QueryParameters.Expand = selectClause is null
                ? ["fields"]
                : [$"fields($select={selectClause})"];

            if (!string.IsNullOrEmpty(filter))
                r.QueryParameters.Filter = filter;
        });

        var rawItems = items?.Value ?? new List<ListItem>();

        var lookupCols = string.IsNullOrEmpty(expandLookups)
            ? Array.Empty<string>()
            : expandLookups.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(c => c.EndsWith("LookupId", StringComparison.OrdinalIgnoreCase) ? c[..^"LookupId".Length] : c)
                .ToArray();

        Dictionary<int, Dictionary<string, object?>>? userMap = null;
        if (lookupCols.Length > 0)
        {
            var ids = new HashSet<int>();
            foreach (var it in rawItems)
            {
                if (it.Fields?.AdditionalData == null) continue;
                foreach (var col in lookupCols)
                {
                    if (it.Fields.AdditionalData.TryGetValue(col + "LookupId", out var v) && v != null
                        && int.TryParse(v.ToString(), out var id))
                        ids.Add(id);
                }
            }
            if (ids.Count > 0)
                userMap = await ResolveUserInfoAsync(client, siteId, ids);
        }

        return rawItems.Select(i =>
        {
            var fld = i.Fields?.AdditionalData != null
                ? i.Fields.AdditionalData.ToDictionary(kv => kv.Key, kv => Normalize(kv.Value))
                : new Dictionary<string, object?>();

            if (userMap != null)
            {
                foreach (var col in lookupCols)
                {
                    if (fld.TryGetValue(col + "LookupId", out var v) && v != null
                        && int.TryParse(v.ToString(), out var id)
                        && userMap.TryGetValue(id, out var info))
                    {
                        fld[col] = info;
                    }
                }
            }

            return new
            {
                i.Id,
                i.WebUrl,
                i.CreatedDateTime,
                i.LastModifiedDateTime,
                CreatedBy = i.CreatedBy?.User?.DisplayName,
                LastModifiedBy = i.LastModifiedBy?.User?.DisplayName,
                Fields = (object)fld
            };
        }).ToList<object>();
    }

    public static async Task<object> ColumnsAsync(string site, string listId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await PageService.ResolveSiteIdAsync(client, site);

        var columns = await client.Sites[siteId].Lists[listId].Columns.GetAsync();

        return columns?.Value?.Select(c => new
        {
            c.DisplayName,
            c.Name,
            c.Description,
            Type = ColumnType(c),
            c.Hidden,
            c.ReadOnly,
            Choices = c.Choice?.Choices
        }).ToList() ?? [];
    }

    private static string? ColumnType(ColumnDefinition c) => c switch
    {
        { Text: not null } => "text",
        { Choice: not null } => "choice",
        { Number: not null } => "number",
        { DateTime: not null } => "dateTime",
        { Boolean: not null } => "boolean",
        { Currency: not null } => "currency",
        { Lookup: not null } => "lookup",
        { PersonOrGroup: not null } => "personOrGroup",
        { HyperlinkOrPicture: not null } => "hyperlinkOrPicture",
        { Calculated: not null } => "calculated",
        { Geolocation: not null } => "geolocation",
        _ => null
    };

    // Graph returns complex list-field values (multi-value choice, lookups) as
    // Kiota UntypedNode instances. System.Text.Json can't serialize those — a
    // multi-choice column would emit as an empty {}. Unwrap them into plain CLR
    // values so arrays/objects serialize like every other field type.
    private static object? Normalize(object? value) => value switch
    {
        UntypedString s => s.GetValue(),
        UntypedBoolean b => b.GetValue(),
        UntypedInteger i => i.GetValue(),
        UntypedLong l => l.GetValue(),
        UntypedDecimal m => m.GetValue(),
        UntypedDouble d => d.GetValue(),
        UntypedFloat f => f.GetValue(),
        UntypedNull => null,
        UntypedArray arr => arr.GetValue().Select(Normalize).ToList(),
        UntypedObject obj => obj.GetValue().ToDictionary(kv => kv.Key, kv => Normalize(kv.Value)),
        _ => value
    };

    private static async Task<Dictionary<int, Dictionary<string, object?>>> ResolveUserInfoAsync(
        GraphServiceClient client, string siteId, HashSet<int> ids)
    {
        var tasks = ids.Select(async id =>
        {
            try
            {
                var item = await client.Sites[siteId].Lists["User Information List"].Items[id.ToString()].GetAsync(r =>
                {
                    r.QueryParameters.Expand = ["fields($select=Title,Name,EMail)"];
                });
                var data = item?.Fields?.AdditionalData;
                return (id, info: new Dictionary<string, object?>
                {
                    ["LookupId"] = id,
                    ["LookupValue"] = data != null && data.TryGetValue("Title", out var t) ? t : null,
                    ["Email"] = data != null && data.TryGetValue("EMail", out var e) ? e : null,
                    ["Name"] = data != null && data.TryGetValue("Name", out var n) ? n : null
                });
            }
            catch
            {
                return (id, info: (Dictionary<string, object?>?)null);
            }
        });
        var results = await Task.WhenAll(tasks);
        return results
            .Where(r => r.info != null)
            .ToDictionary(r => r.id, r => r.info!);
    }

    private static string? BuildFieldsSelect(string? fields, string? expandLookups)
    {
        var lookupCols = string.IsNullOrEmpty(expandLookups)
            ? Array.Empty<string>()
            : expandLookups.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(c => c.EndsWith("LookupId", StringComparison.OrdinalIgnoreCase) ? c[..^"LookupId".Length] : c)
                .ToArray();

        // No explicit --fields → leave $select clauseless so Graph returns all
        // default fields. The User-Info-List join in ItemsAsync still fires
        // because every *LookupId is present in the default response; we just
        // don't need to mention them here.
        if (string.IsNullOrEmpty(fields))
            return null;

        var parts = new List<string>(
            fields.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

        // For lookup columns, ask Graph for the integer *LookupId form (the
        // raw stored value). The User Information List join in ItemsAsync
        // then resolves it to a {LookupId, Email, Name, Title} object that
        // gets emitted under the bare column name. Selecting the bare name
        // would make Graph return only a display-name string with no ID,
        // leaving the join logic with nothing to look up.
        var expanded = new HashSet<string>(lookupCols, StringComparer.OrdinalIgnoreCase);
        parts.RemoveAll(p => expanded.Contains(p));
        foreach (var col in lookupCols)
        {
            var lookupIdCol = col + "LookupId";
            if (!parts.Contains(lookupIdCol, StringComparer.OrdinalIgnoreCase))
                parts.Add(lookupIdCol);
        }

        return string.Join(",", parts);
    }
}
