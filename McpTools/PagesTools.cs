using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class PagesTools
{
    [McpServerTool(Name = "pages_list"), Description("List pages on a SharePoint site")]
    public static async Task<string> List(
        [Description("SharePoint site ID or hostname (e.g. contoso.sharepoint.com:/sites/team)")] string site,
        [Description("Number of results (default: 25)")] int top = 25,
        [Description("Search pages by name or title")] string? search = null)
    {
        try
        {
            var result = await PageService.ListAsync(site, top, search);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_get"), Description("Get page details and optionally its content")]
    public static async Task<string> Get(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Include full page canvas layout content")] bool expandContent = false)
    {
        try
        {
            var result = await PageService.GetAsync(site, pageId, expandContent);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_create"), Description("Create a new SharePoint page")]
    public static async Task<string> Create(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page title")] string title,
        [Description("Page file name (e.g. my-page.aspx); auto-generated from title if omitted")] string? name = null,
        [Description("HTML content for a simple single-section text page")] string? content = null,
        [Description("Full canvas layout JSON for complex multi-section pages (overrides content)")] string? canvasLayoutJson = null,
        [Description("Publish the page immediately after creation")] bool publish = false)
    {
        try
        {
            var result = await PageService.CreateAsync(site, title, name, content, canvasLayoutJson, publish);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_update"), Description("Update an existing SharePoint page")]
    public static async Task<string> Update(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("New page title")] string? title = null,
        [Description("HTML content to replace page body (single-section)")] string? content = null,
        [Description("Full canvas layout JSON for complex multi-section pages (overrides content)")] string? canvasLayoutJson = null,
        [Description("Publish the page after updating")] bool publish = false)
    {
        try
        {
            var result = await PageService.UpdateAsync(site, pageId, title, content, canvasLayoutJson, publish);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_publish"), Description("Publish a SharePoint page")]
    public static async Task<string> Publish(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId)
    {
        try
        {
            var result = await PageService.PublishAsync(site, pageId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    // ── Sections ──────────────────────────────────────────────────────

    [McpServerTool(Name = "pages_sections_list"), Description("List sections on a SharePoint page")]
    public static async Task<string> SectionsList(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId)
    {
        try
        {
            var result = await PageService.ListSectionsAsync(site, pageId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_sections_get"), Description("Get a section with its columns and webparts")]
    public static async Task<string> SectionsGet(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID")] string sectionId)
    {
        try
        {
            var result = await PageService.GetSectionAsync(site, pageId, sectionId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_sections_create"), Description("Add a new section to a SharePoint page")]
    public static async Task<string> SectionsCreate(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section layout: fullWidth, oneColumn, twoColumn, threeColumn, oneThirdLeftColumn, oneThirdRightColumn")] string layout,
        [Description("Section emphasis: none, neutral, soft, strong")] string? emphasis = null)
    {
        try
        {
            var result = await PageService.CreateSectionAsync(site, pageId, layout, emphasis);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_sections_update"), Description("Update a section on a SharePoint page")]
    public static async Task<string> SectionsUpdate(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID")] string sectionId,
        [Description("New section layout")] string? layout = null,
        [Description("New section emphasis")] string? emphasis = null)
    {
        try
        {
            var result = await PageService.UpdateSectionAsync(site, pageId, sectionId, layout, emphasis);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_sections_delete"), Description("Delete a section from a SharePoint page")]
    public static async Task<string> SectionsDelete(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID")] string sectionId)
    {
        try
        {
            var result = await PageService.DeleteSectionAsync(site, pageId, sectionId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    // ── Columns ──────────────────────────────────────────────────────

    [McpServerTool(Name = "pages_columns_list"), Description("List columns in a page section with their webparts")]
    public static async Task<string> ColumnsList(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID")] string sectionId)
    {
        try
        {
            var result = await PageService.ListColumnsAsync(site, pageId, sectionId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    // ── WebParts ─────────────────────────────────────────────────────

    [McpServerTool(Name = "pages_webparts_list"), Description("List webparts on a page, optionally filtered by section and column")]
    public static async Task<string> WebPartsList(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID (optional, requires column-id)")] string? sectionId = null,
        [Description("Column ID (optional, requires section-id)")] string? columnId = null)
    {
        try
        {
            var result = await PageService.ListWebPartsAsync(site, pageId, sectionId, columnId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_webparts_get"), Description("Get a specific webpart")]
    public static async Task<string> WebPartsGet(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID")] string sectionId,
        [Description("Column ID")] string columnId,
        [Description("WebPart ID")] string webpartId)
    {
        try
        {
            var result = await PageService.GetWebPartAsync(site, pageId, sectionId, columnId, webpartId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_webparts_create"), Description("Add a webpart to a column on a SharePoint page")]
    public static async Task<string> WebPartsCreate(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID")] string sectionId,
        [Description("Column ID")] string columnId,
        [Description("HTML content (creates a TextWebPart)")] string? innerHtml = null,
        [Description("Standard webpart type GUID (creates a StandardWebPart)")] string? webPartType = null,
        [Description("JSON data for standard webpart configuration")] string? dataJson = null)
    {
        try
        {
            var result = await PageService.CreateWebPartAsync(site, pageId, sectionId, columnId, innerHtml, webPartType, dataJson);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_webparts_update"), Description("Update a webpart on a SharePoint page")]
    public static async Task<string> WebPartsUpdate(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID")] string sectionId,
        [Description("Column ID")] string columnId,
        [Description("WebPart ID")] string webpartId,
        [Description("New HTML content (for TextWebPart)")] string? innerHtml = null,
        [Description("New JSON data (for StandardWebPart)")] string? dataJson = null)
    {
        try
        {
            var result = await PageService.UpdateWebPartAsync(site, pageId, sectionId, columnId, webpartId, innerHtml, dataJson);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "pages_webparts_delete"), Description("Delete a webpart from a SharePoint page")]
    public static async Task<string> WebPartsDelete(
        [Description("SharePoint site ID or hostname")] string site,
        [Description("Page ID")] string pageId,
        [Description("Section ID")] string sectionId,
        [Description("Column ID")] string columnId,
        [Description("WebPart ID")] string webpartId)
    {
        try
        {
            var result = await PageService.DeleteWebPartAsync(site, pageId, sectionId, columnId, webpartId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
