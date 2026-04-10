using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions.Serialization;

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

    public static async Task<object> CreateAsync(
        string site, string title, string? name, string? content, string? canvasLayoutJson, bool publish)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        var pageName = name ?? SlugifyTitle(title) + ".aspx";
        if (!pageName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
            pageName += ".aspx";

        var sitePage = new SitePage
        {
            OdataType = "#microsoft.graph.sitePage",
            Title = title,
            Name = pageName,
            PageLayout = PageLayoutType.Article
        };

        if (!string.IsNullOrEmpty(canvasLayoutJson))
            sitePage.CanvasLayout = ParseCanvasLayoutJson(canvasLayoutJson);
        else if (!string.IsNullOrEmpty(content))
            sitePage.CanvasLayout = BuildSingleSectionLayout(content);

        var created = await client.Sites[siteId].Pages.PostAsync(sitePage);

        var published = false;
        if (publish && created?.Id != null)
        {
            await PublishInternalAsync(client, siteId, created.Id);
            published = true;
        }

        return new
        {
            status = "created",
            id = created?.Id,
            title,
            name = pageName,
            webUrl = created?.WebUrl,
            published
        };
    }

    public static async Task<object> UpdateAsync(
        string site, string pageId, string? title, string? content, string? canvasLayoutJson, bool publish)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        var sitePage = new SitePage
        {
            OdataType = "#microsoft.graph.sitePage"
        };

        if (title != null) sitePage.Title = title;

        if (!string.IsNullOrEmpty(canvasLayoutJson))
            sitePage.CanvasLayout = ParseCanvasLayoutJson(canvasLayoutJson);
        else if (!string.IsNullOrEmpty(content))
            sitePage.CanvasLayout = BuildSingleSectionLayout(content);

        var updated = await client.Sites[siteId].Pages[pageId].PatchAsync(sitePage);

        var published = false;
        if (publish)
        {
            await PublishInternalAsync(client, siteId, pageId);
            published = true;
        }

        return new
        {
            status = "updated",
            id = updated?.Id ?? pageId,
            published
        };
    }

    public static async Task<object> PublishAsync(string site, string pageId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);
        await PublishInternalAsync(client, siteId, pageId);
        return new { status = "published", id = pageId };
    }

    private static async Task PublishInternalAsync(GraphServiceClient client, string siteId, string pageId)
    {
        var requestInfo = new Microsoft.Kiota.Abstractions.RequestInformation
        {
            HttpMethod = Microsoft.Kiota.Abstractions.Method.POST,
            UrlTemplate = $"https://graph.microsoft.com/v1.0/sites/{siteId}/pages/{pageId}/microsoft.graph.sitePage/publish"
        };
        await client.RequestAdapter.SendNoContentAsync(requestInfo);
    }

    private static CanvasLayout BuildSingleSectionLayout(string htmlContent)
    {
        return new CanvasLayout
        {
            HorizontalSections =
            [
                new HorizontalSection
                {
                    Layout = HorizontalSectionLayoutType.FullWidth,
                    Id = "1",
                    Emphasis = SectionEmphasisType.None,
                    Columns =
                    [
                        new HorizontalSectionColumn
                        {
                            Id = "1",
                            Width = 12,
                            Webparts =
                            [
                                new TextWebPart
                                {
                                    OdataType = "#microsoft.graph.textWebPart",
                                    InnerHtml = htmlContent
                                }
                            ]
                        }
                    ]
                }
            ]
        };
    }

    private static CanvasLayout ParseCanvasLayoutJson(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        var layout = new CanvasLayout();

        if (root.TryGetProperty("horizontalSections", out var sectionsEl))
        {
            layout.HorizontalSections = [];
            foreach (var sectionEl in sectionsEl.EnumerateArray())
            {
                var section = new HorizontalSection
                {
                    Id = sectionEl.TryGetProperty("id", out var idEl) ? idEl.GetString() : null,
                    Layout = ParseEnum<HorizontalSectionLayoutType>(sectionEl, "layout"),
                    Emphasis = ParseEnum<SectionEmphasisType>(sectionEl, "emphasis")
                };

                if (sectionEl.TryGetProperty("columns", out var columnsEl))
                {
                    section.Columns = [];
                    foreach (var colEl in columnsEl.EnumerateArray())
                    {
                        var column = new HorizontalSectionColumn
                        {
                            Id = colEl.TryGetProperty("id", out var colIdEl) ? colIdEl.GetString() : null,
                            Width = colEl.TryGetProperty("width", out var widthEl) ? widthEl.GetInt32() : null
                        };

                        if (colEl.TryGetProperty("webparts", out var webpartsEl))
                        {
                            column.Webparts = [];
                            foreach (var wpEl in webpartsEl.EnumerateArray())
                            {
                                column.Webparts.Add(ParseWebPart(wpEl));
                            }
                        }

                        section.Columns.Add(column);
                    }
                }

                layout.HorizontalSections.Add(section);
            }
        }

        return layout;
    }

    private static WebPart ParseWebPart(JsonElement el)
    {
        if (el.TryGetProperty("webPartType", out _))
        {
            var wp = new StandardWebPart
            {
                OdataType = "#microsoft.graph.standardWebPart",
                Id = el.TryGetProperty("id", out var idEl) ? idEl.GetString() : null,
                WebPartType = el.TryGetProperty("webPartType", out var typeEl) ? typeEl.GetString() : null
            };

            if (el.TryGetProperty("data", out var dataEl))
            {
                wp.Data = new WebPartData();
                wp.Data.AdditionalData = JsonToAdditionalData(dataEl);
            }

            return wp;
        }
        else
        {
            var wp = new TextWebPart
            {
                OdataType = "#microsoft.graph.textWebPart",
                Id = el.TryGetProperty("id", out var idEl) ? idEl.GetString() : null,
                InnerHtml = el.TryGetProperty("innerHtml", out var htmlEl) ? htmlEl.GetString() : null
            };
            return wp;
        }
    }

    private static Dictionary<string, object> JsonToAdditionalData(JsonElement el)
    {
        var dict = new Dictionary<string, object>();
        foreach (var prop in el.EnumerateObject())
        {
            dict[prop.Name] = prop.Value.ValueKind switch
            {
                JsonValueKind.String => prop.Value.GetString()!,
                JsonValueKind.Number => prop.Value.GetRawText(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Object => new UntypedObject(
                    prop.Value.EnumerateObject().ToDictionary(
                        p => p.Name,
                        p => (UntypedNode)new UntypedString(p.Value.ToString()))),
                JsonValueKind.Array => new UntypedArray(
                    prop.Value.EnumerateArray().Select(
                        item => (UntypedNode)new UntypedString(item.ToString())).ToList()),
                _ => prop.Value.GetRawText()
            };
        }
        return dict;
    }

    private static T? ParseEnum<T>(JsonElement el, string propertyName) where T : struct, Enum
    {
        if (!el.TryGetProperty(propertyName, out var val)) return null;
        var str = val.GetString();
        if (string.IsNullOrEmpty(str)) return null;
        return Enum.TryParse<T>(str, ignoreCase: true, out var result) ? result : null;
    }

    private static string SlugifyTitle(string title)
    {
        var slug = title.ToLowerInvariant();
        slug = Regex.Replace(slug, @"[^a-z0-9\s-]", "");
        slug = Regex.Replace(slug, @"[\s]+", "-");
        slug = Regex.Replace(slug, @"-{2,}", "-");
        return slug.Trim('-');
    }

    // ── Sections ──────────────────────────────────────────────────────

    public static async Task<object> ListSectionsAsync(string site, string pageId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        var response = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections.GetAsync(r =>
            {
                r.QueryParameters.Expand = ["columns"];
            });

        var sections = response?.Value ?? [];
        return sections.Select(s => new
        {
            s.Id,
            Layout = s.Layout?.ToString(),
            Emphasis = s.Emphasis?.ToString(),
            Columns = s.Columns?.Select(c => new
            {
                c.Id,
                c.Width,
                WebpartCount = c.Webparts?.Count ?? 0
            }).ToList()
        }).ToList();
    }

    public static async Task<object> GetSectionAsync(string site, string pageId, string sectionId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        var section = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections[sectionId].GetAsync(r =>
            {
                r.QueryParameters.Expand = ["columns($expand=webparts)"];
            });

        return new
        {
            section?.Id,
            Layout = section?.Layout?.ToString(),
            Emphasis = section?.Emphasis?.ToString(),
            Columns = section?.Columns?.Select(c => new
            {
                c.Id,
                c.Width,
                Webparts = c.Webparts?.Select(FormatWebPart).ToList()
            }).ToList()
        };
    }

    public static async Task<object> CreateSectionAsync(
        string site, string pageId, string layout, string? emphasis)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        var section = new HorizontalSection
        {
            Layout = Enum.Parse<HorizontalSectionLayoutType>(layout, ignoreCase: true),
            Emphasis = !string.IsNullOrEmpty(emphasis)
                ? Enum.Parse<SectionEmphasisType>(emphasis, ignoreCase: true)
                : SectionEmphasisType.None
        };

        var created = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections.PostAsync(section);

        return new
        {
            status = "created",
            id = created?.Id,
            Layout = created?.Layout?.ToString(),
            Emphasis = created?.Emphasis?.ToString()
        };
    }

    public static async Task<object> UpdateSectionAsync(
        string site, string pageId, string sectionId, string? layout, string? emphasis)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        var section = new HorizontalSection();
        if (layout != null) section.Layout = Enum.Parse<HorizontalSectionLayoutType>(layout, ignoreCase: true);
        if (emphasis != null) section.Emphasis = Enum.Parse<SectionEmphasisType>(emphasis, ignoreCase: true);

        var updated = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections[sectionId].PatchAsync(section);

        return new { status = "updated", id = updated?.Id ?? sectionId };
    }

    public static async Task<object> DeleteSectionAsync(string site, string pageId, string sectionId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections[sectionId].DeleteAsync();

        return new { status = "deleted", id = sectionId };
    }

    // ── Columns ──────────────────────────────────────────────────────

    public static async Task<object> ListColumnsAsync(string site, string pageId, string sectionId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        var response = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections[sectionId].Columns.GetAsync(r =>
            {
                r.QueryParameters.Expand = ["webparts"];
            });

        var columns = response?.Value ?? [];
        return columns.Select(c => new
        {
            c.Id,
            c.Width,
            Webparts = c.Webparts?.Select(FormatWebPart).ToList()
        }).ToList();
    }

    // ── WebParts ─────────────────────────────────────────────────────

    public static async Task<object> ListWebPartsAsync(
        string site, string pageId, string? sectionId, string? columnId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        List<WebPart> webparts;
        if (!string.IsNullOrEmpty(sectionId) && !string.IsNullOrEmpty(columnId))
        {
            var response = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
                .HorizontalSections[sectionId].Columns[columnId].Webparts.GetAsync();
            webparts = response?.Value ?? [];
        }
        else
        {
            var response = await client.Sites[siteId].Pages[pageId].GraphSitePage.WebParts.GetAsync();
            webparts = response?.Value ?? [];
        }

        return webparts.Select(FormatWebPart).ToList();
    }

    public static async Task<object> GetWebPartAsync(
        string site, string pageId, string sectionId, string columnId, string webpartId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        var wp = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections[sectionId].Columns[columnId].Webparts[webpartId].GetAsync();

        return FormatWebPart(wp!);
    }

    public static async Task<object> CreateWebPartAsync(
        string site, string pageId, string sectionId, string columnId,
        string? innerHtml, string? webPartType, string? dataJson)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        WebPart webpart;
        if (!string.IsNullOrEmpty(innerHtml))
        {
            webpart = new TextWebPart
            {
                OdataType = "#microsoft.graph.textWebPart",
                InnerHtml = innerHtml
            };
        }
        else if (!string.IsNullOrEmpty(webPartType))
        {
            var wp = new StandardWebPart
            {
                OdataType = "#microsoft.graph.standardWebPart",
                WebPartType = webPartType
            };
            if (!string.IsNullOrEmpty(dataJson))
            {
                wp.Data = new WebPartData();
                using var doc = JsonDocument.Parse(dataJson);
                wp.Data.AdditionalData = JsonToAdditionalData(doc.RootElement);
            }
            webpart = wp;
        }
        else
        {
            throw new InvalidOperationException("Either --inner-html or --webpart-type must be provided.");
        }

        var created = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections[sectionId].Columns[columnId].Webparts.PostAsync(webpart);

        return new { status = "created", id = created?.Id, type = created?.OdataType };
    }

    public static async Task<object> UpdateWebPartAsync(
        string site, string pageId, string sectionId, string columnId, string webpartId,
        string? innerHtml, string? dataJson)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        WebPart webpart;
        if (!string.IsNullOrEmpty(innerHtml))
        {
            webpart = new TextWebPart
            {
                OdataType = "#microsoft.graph.textWebPart",
                InnerHtml = innerHtml
            };
        }
        else if (!string.IsNullOrEmpty(dataJson))
        {
            var wp = new StandardWebPart
            {
                OdataType = "#microsoft.graph.standardWebPart"
            };
            wp.Data = new WebPartData();
            using var doc = JsonDocument.Parse(dataJson);
            wp.Data.AdditionalData = JsonToAdditionalData(doc.RootElement);
            webpart = wp;
        }
        else
        {
            throw new InvalidOperationException("Either --inner-html or --data-json must be provided.");
        }

        var updated = await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections[sectionId].Columns[columnId].Webparts[webpartId].PatchAsync(webpart);

        return new { status = "updated", id = updated?.Id ?? webpartId };
    }

    public static async Task<object> DeleteWebPartAsync(
        string site, string pageId, string sectionId, string columnId, string webpartId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var siteId = await ResolveSiteIdAsync(client, site);

        await client.Sites[siteId].Pages[pageId].GraphSitePage.CanvasLayout
            .HorizontalSections[sectionId].Columns[columnId].Webparts[webpartId].DeleteAsync();

        return new { status = "deleted", id = webpartId };
    }

    private static object FormatWebPart(WebPart wp)
    {
        return new
        {
            wp.Id,
            ODataType = wp.OdataType,
            InnerHtml = (wp as TextWebPart)?.InnerHtml,
            WebPartType = (wp as StandardWebPart)?.WebPartType,
            Data = (wp as StandardWebPart)?.Data?.AdditionalData
        };
    }

    // ── Helpers ──────────────────────────────────────────────────────

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
