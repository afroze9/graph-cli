using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Services;

public static class PageService
{
    public static async Task<object> ListAsync(string site, int top, string? search)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        List<BaseSitePage> items;

        if (!string.IsNullOrEmpty(search))
        {
            items = [];
            var response = await client.Sites[siteId].Pages
                .GetAsync(r =>
                {
                    r.QueryParameters.Top = 100;
                    r.QueryParameters.Select = ["id", "name", "title", "webUrl",
                        "createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy"];
                });

            while (response?.Value != null)
            {
                items.AddRange(response.Value.Where(p =>
                    (p.Name?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    ((p as SitePage)?.Title?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false)));

                if (items.Count >= top || string.IsNullOrEmpty(response.OdataNextLink))
                    break;

                response = await client.Sites[siteId].Pages
                    .WithUrl(response.OdataNextLink)
                    .GetAsync();
            }

            items = items.Take(top).ToList();
        }
        else
        {
            var response = await client.Sites[siteId].Pages
                .GetAsync(r =>
                {
                    r.QueryParameters.Top = top;
                    r.QueryParameters.Select = ["id", "name", "title", "webUrl",
                        "createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy"];
                    r.QueryParameters.Orderby = ["lastModifiedDateTime desc"];
                });

            items = response?.Value ?? [];
        }

        return items.Select(p => new
        {
            p.Id,
            Title = (p as SitePage)?.Title ?? p.Name,
            p.Name,
            PageLayout = (p as SitePage)?.PageLayout?.ToString(),
            PublishingState = (p as SitePage)?.PublishingState?.Level?.ToString(),
            p.WebUrl,
            p.CreatedDateTime,
            p.LastModifiedDateTime,
            CreatedBy = p.CreatedBy?.User?.DisplayName,
            LastModifiedBy = p.LastModifiedBy?.User?.DisplayName
        }).ToList();
    }

    public static async Task<object> GetAsync(string site, string pageId, bool expandContent)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        // Try sitePage cast first for rich properties + canvas, fall back to base
        SitePage? page = null;
        try
        {
            page = await client.Sites[siteId].Pages[pageId].GraphSitePage
                .GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id", "name", "title", "webUrl", "pageLayout", "promotionKind",
                        "createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy",
                        "publishingState", "titleArea", "description", "thumbnailWebUrl",
                        "showComments", "showRecommendedPages"];
                    if (expandContent)
                    {
                        r.QueryParameters.Expand = ["canvasLayout"];
                    }
                });
        }
        catch (ODataError) { /* page may not support sitePage cast */ }

        if (page != null)
        {
            object? canvasLayout = null;
            if (expandContent && page.CanvasLayout != null)
            {
                canvasLayout = page.CanvasLayout.HorizontalSections?.Select(s => new
                {
                    s.Id,
                    Layout = s.Layout?.ToString(),
                    Columns = s.Columns?.Select(c => new
                    {
                        c.Id,
                        c.Width,
                        Webparts = c.Webparts?.Select(wp => new
                        {
                            ODataType = wp.OdataType,
                            Id = wp.Id,
                            InnerHtml = (wp as TextWebPart)?.InnerHtml,
                            WebPartType = (wp as StandardWebPart)?.WebPartType,
                        }).ToList()
                    }).ToList()
                }).ToList();
            }

            return new
            {
                page.Id,
                page.Title,
                page.Name,
                page.Description,
                PageLayout = page.PageLayout?.ToString(),
                PromotionKind = page.PromotionKind?.ToString(),
                PublishingState = page.PublishingState?.Level?.ToString(),
                PublishingVersion = page.PublishingState?.VersionId,
                page.ShowComments,
                page.ShowRecommendedPages,
                page.WebUrl,
                page.ThumbnailWebUrl,
                page.CreatedDateTime,
                page.LastModifiedDateTime,
                CreatedBy = page.CreatedBy?.User?.DisplayName,
                LastModifiedBy = page.LastModifiedBy?.User?.DisplayName,
                TitleArea = page.TitleArea != null ? new
                {
                    Layout = page.TitleArea.Layout?.ToString(),
                    TextAlignment = page.TitleArea.TextAlignment?.ToString(),
                    ShowAuthor = page.TitleArea.ShowAuthor,
                    ShowPublishedDate = page.TitleArea.ShowPublishedDate,
                    page.TitleArea.ImageWebUrl,
                    page.TitleArea.EnableGradientEffect
                } : null,
                CanvasLayout = canvasLayout
            };
        }
        else
        {
            // Fall back to base page (no canvas support)
            var basePage = await client.Sites[siteId].Pages[pageId].GetAsync();

            if (basePage == null)
                throw new InvalidOperationException("Page not found.");

            return new
            {
                basePage.Id,
                Title = (basePage as SitePage)?.Title ?? basePage.Name,
                basePage.Name,
                basePage.WebUrl,
                basePage.CreatedDateTime,
                basePage.LastModifiedDateTime,
                CreatedBy = basePage.CreatedBy?.User?.DisplayName,
                LastModifiedBy = basePage.LastModifiedBy?.User?.DisplayName
            };
        }
    }

    /// <summary>
    /// Resolves a site identifier to a site ID. Accepts:
    /// - Full hostname path: "contoso.sharepoint.com:/sites/TSASite"
    /// - Graph site ID: "contoso.sharepoint.com,guid,guid"
    /// - Bare site name: "TSASite" (resolved via site cache)
    /// </summary>
    internal static async Task<string> ResolveSiteIdAsync(
        GraphServiceClient client, string site, CancellationToken ct = default)
    {
        var resolved = site;

        // If it doesn't contain a dot or comma, treat it as a bare name and check cache
        if (!site.Contains('.') && !site.Contains(','))
        {
            var cached = SiteCacheService.Resolve(site);
            if (cached != null)
            {
                resolved = cached;
            }
            else
            {
                throw new InvalidOperationException(
                    $"Could not resolve site '{site}' from cache. " +
                    $"Run 'graph-cli sites search \"{site}\"' first to populate the cache, " +
                    $"or use a full hostname path like 'host.sharepoint.com:/sites/{site}'.");
            }
        }

        var siteObj = await client.Sites[resolved].GetAsync(r =>
        {
            r.QueryParameters.Select = ["id", "name", "displayName", "webUrl"];
        }, ct);

        if (siteObj?.Id == null)
            throw new InvalidOperationException($"Could not resolve site: {site}");

        // Cache the resolved site
        if (siteObj.Name != null)
            SiteCacheService.Upsert(siteObj.Id, siteObj.Name, siteObj.DisplayName, siteObj.WebUrl);

        return siteObj.Id;
    }
}
