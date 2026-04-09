using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class PagesCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var pagesCommand = new Command("pages", "SharePoint site page operations");

        pagesCommand.Subcommands.Add(BuildList(formatOption));
        pagesCommand.Subcommands.Add(BuildGet(formatOption));

        return pagesCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname (e.g. contoso.sharepoint.com:/sites/team)", Required = true };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of results to return" };
        var searchOption = new Option<string?>("--search") { Description = "Search pages by name or title (client-side, fetches all pages)" };
        var cmd = new Command("list", "List pages on a SharePoint site") { siteOption, topOption, searchOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var site = parseResult.GetValue(siteOption)!;
            var top = parseResult.GetValue(topOption);
            var search = parseResult.GetValue(searchOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var siteId = await ResolveSiteIdAsync(client, site, ct);

                List<BaseSitePage> items;

                if (!string.IsNullOrEmpty(search))
                {
                    // Paginate through all pages, filter client-side
                    items = [];
                    var response = await client.Sites[siteId].Pages
                        .GetAsync(r =>
                        {
                            r.QueryParameters.Top = 100;
                            r.QueryParameters.Select = ["id", "name", "title", "webUrl",
                                "createdDateTime", "lastModifiedDateTime", "createdBy", "lastModifiedBy"];
                        }, ct);

                    while (response?.Value != null)
                    {
                        items.AddRange(response.Value.Where(p =>
                            (p.Name?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false) ||
                            ((p as SitePage)?.Title?.Contains(search, StringComparison.OrdinalIgnoreCase) ?? false)));

                        if (items.Count >= top || string.IsNullOrEmpty(response.OdataNextLink))
                            break;

                        response = await client.Sites[siteId].Pages
                            .WithUrl(response.OdataNextLink)
                            .GetAsync(cancellationToken: ct);
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
                        }, ct);

                    items = response?.Value ?? [];
                }

                var results = items.Select(p => new
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
        var pageArg = new Argument<string>("page-id") { Description = "Page ID" };
        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname", Required = true };
        var expandContentOption = new Option<bool>("--expand-content") { Description = "Include full page canvas layout content" };
        var cmd = new Command("get", "Get page details and optionally its content") { pageArg, siteOption, expandContentOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var pageId = parseResult.GetValue(pageArg)!;
            var site = parseResult.GetValue(siteOption)!;
            var expandContent = parseResult.GetValue(expandContentOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var siteId = await ResolveSiteIdAsync(client, site, ct);

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
                        }, ct);
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

                    OutputService.Print(new
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
                    }, format);
                }
                else
                {
                    // Fall back to base page (no canvas support)
                    var basePage = await client.Sites[siteId].Pages[pageId]
                        .GetAsync(cancellationToken: ct);

                    if (basePage == null)
                    {
                        OutputService.PrintError("not_found", "Page not found.");
                        Environment.ExitCode = 1;
                        return;
                    }

                    OutputService.Print(new
                    {
                        basePage.Id,
                        Title = (basePage as SitePage)?.Title ?? basePage.Name,
                        basePage.Name,
                        basePage.WebUrl,
                        basePage.CreatedDateTime,
                        basePage.LastModifiedDateTime,
                        CreatedBy = basePage.CreatedBy?.User?.DisplayName,
                        LastModifiedBy = basePage.LastModifiedBy?.User?.DisplayName
                    }, format);
                }
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    /// <summary>
    /// Resolves a site identifier to a site ID. Accepts:
    /// - Full hostname path: "contoso.sharepoint.com:/sites/TSASite"
    /// - Graph site ID: "contoso.sharepoint.com,guid,guid"
    /// - Bare site name: "TSASite" (resolved via site cache)
    /// </summary>
    internal static async Task<string> ResolveSiteIdAsync(
        GraphServiceClient client, string site, CancellationToken ct)
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
