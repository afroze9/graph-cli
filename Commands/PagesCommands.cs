using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class PagesCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var pagesCommand = new Command("pages", "SharePoint site page operations");

        pagesCommand.Subcommands.Add(BuildList(formatOption));
        pagesCommand.Subcommands.Add(BuildGet(formatOption));
        pagesCommand.Subcommands.Add(BuildCreate(formatOption));
        pagesCommand.Subcommands.Add(BuildUpdate(formatOption));
        pagesCommand.Subcommands.Add(BuildPublish(formatOption));
        pagesCommand.Subcommands.Add(BuildSections(formatOption));
        pagesCommand.Subcommands.Add(BuildColumns(formatOption));
        pagesCommand.Subcommands.Add(BuildWebParts(formatOption));

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
                var result = await PageService.ListAsync(site, top, search);
                OutputService.Print(result, format);
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
                var result = await PageService.GetAsync(site, pageId, expandContent);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildCreate(Option<string> formatOption)
    {
        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname (e.g. contoso.sharepoint.com:/sites/team)", Required = true };
        var titleOption = new Option<string>("--title") { Description = "Page title", Required = true };
        var nameOption = new Option<string?>("--name") { Description = "Page file name (e.g. my-page.aspx); auto-generated from title if omitted" };
        var contentOption = new Option<string?>("--content") { Description = "HTML content for a single-section text page" };
        var publishOption = new Option<bool>("--publish") { Description = "Publish the page immediately after creation" };
        var cmd = new Command("create", "Create a new SharePoint page") { siteOption, titleOption, nameOption, contentOption, publishOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var site = parseResult.GetValue(siteOption)!;
            var title = parseResult.GetValue(titleOption)!;
            var name = parseResult.GetValue(nameOption);
            var content = parseResult.GetValue(contentOption);
            var publish = parseResult.GetValue(publishOption);
            try
            {
                var result = await PageService.CreateAsync(site, title, name, content, canvasLayoutJson: null, publish);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildUpdate(Option<string> formatOption)
    {
        var pageArg = new Argument<string>("page-id") { Description = "Page ID" };
        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname", Required = true };
        var titleOption = new Option<string?>("--title") { Description = "New page title" };
        var contentOption = new Option<string?>("--content") { Description = "HTML content to replace page body (single-section)" };
        var publishOption = new Option<bool>("--publish") { Description = "Publish the page after updating" };
        var cmd = new Command("update", "Update an existing SharePoint page") { pageArg, siteOption, titleOption, contentOption, publishOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var pageId = parseResult.GetValue(pageArg)!;
            var site = parseResult.GetValue(siteOption)!;
            var title = parseResult.GetValue(titleOption);
            var content = parseResult.GetValue(contentOption);
            var publish = parseResult.GetValue(publishOption);
            try
            {
                var result = await PageService.UpdateAsync(site, pageId, title, content, canvasLayoutJson: null, publish);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildPublish(Option<string> formatOption)
    {
        var pageArg = new Argument<string>("page-id") { Description = "Page ID" };
        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname", Required = true };
        var cmd = new Command("publish", "Publish a SharePoint page") { pageArg, siteOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var pageId = parseResult.GetValue(pageArg)!;
            var site = parseResult.GetValue(siteOption)!;
            try
            {
                var result = await PageService.PublishAsync(site, pageId);
                OutputService.Print(result, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    // ── Sections subcommand group ────────────────────────────────────

    private static Command BuildSections(Option<string> formatOption)
    {
        var group = new Command("sections", "Manage page sections (horizontal layout sections)");

        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname", Required = true };
        var pageIdOption = new Option<string>("--page-id") { Description = "Page ID", Required = true };

        // list
        var listCmd = new Command("list", "List sections on a page") { siteOption, pageIdOption };
        listCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.ListSectionsAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!);
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        // get
        var sectionArg = new Argument<string>("section-id") { Description = "Section ID" };
        var getCmd = new Command("get", "Get a section with its columns and webparts") { sectionArg, siteOption, pageIdOption };
        getCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.GetSectionAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionArg)!);
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        // create
        var layoutOption = new Option<string>("--layout") { Description = "Section layout: fullWidth, oneColumn, twoColumn, threeColumn, oneThirdLeftColumn, oneThirdRightColumn", Required = true };
        var emphasisOption = new Option<string?>("--emphasis") { Description = "Section emphasis: none, neutral, soft, strong" };
        var createCmd = new Command("create", "Add a new section to a page") { siteOption, pageIdOption, layoutOption, emphasisOption };
        createCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.CreateSectionAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(layoutOption)!, parseResult.GetValue(emphasisOption));
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        // update
        var updateLayoutOption = new Option<string?>("--layout") { Description = "New section layout" };
        var updateEmphasisOption = new Option<string?>("--emphasis") { Description = "New section emphasis" };
        var updateCmd = new Command("update", "Update a section") { sectionArg, siteOption, pageIdOption, updateLayoutOption, updateEmphasisOption };
        updateCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.UpdateSectionAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionArg)!, parseResult.GetValue(updateLayoutOption), parseResult.GetValue(updateEmphasisOption));
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        // delete
        var deleteCmd = new Command("delete", "Delete a section") { sectionArg, siteOption, pageIdOption };
        deleteCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.DeleteSectionAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionArg)!);
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        group.Subcommands.Add(listCmd);
        group.Subcommands.Add(getCmd);
        group.Subcommands.Add(createCmd);
        group.Subcommands.Add(updateCmd);
        group.Subcommands.Add(deleteCmd);
        return group;
    }

    // ── Columns subcommand group ─────────────────────────────────────

    private static Command BuildColumns(Option<string> formatOption)
    {
        var group = new Command("columns", "List columns within a page section");

        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname", Required = true };
        var pageIdOption = new Option<string>("--page-id") { Description = "Page ID", Required = true };
        var sectionIdOption = new Option<string>("--section-id") { Description = "Section ID", Required = true };

        var listCmd = new Command("list", "List columns in a section with their webparts") { siteOption, pageIdOption, sectionIdOption };
        listCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.ListColumnsAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionIdOption)!);
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        group.Subcommands.Add(listCmd);
        return group;
    }

    // ── WebParts subcommand group ────────────────────────────────────

    private static Command BuildWebParts(Option<string> formatOption)
    {
        var group = new Command("webparts", "Manage webparts on a page");

        var siteOption = new Option<string>("--site") { Description = "SharePoint site ID or hostname", Required = true };
        var pageIdOption = new Option<string>("--page-id") { Description = "Page ID", Required = true };
        var sectionIdOption = new Option<string>("--section-id") { Description = "Section ID" };
        var sectionIdReqOption = new Option<string>("--section-id") { Description = "Section ID", Required = true };
        var columnIdOption = new Option<string>("--column-id") { Description = "Column ID" };
        var columnIdReqOption = new Option<string>("--column-id") { Description = "Column ID", Required = true };

        // list
        var listCmd = new Command("list", "List webparts (optionally filtered by section and column)") { siteOption, pageIdOption, sectionIdOption, columnIdOption };
        listCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.ListWebPartsAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionIdOption), parseResult.GetValue(columnIdOption));
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        // get
        var webpartArg = new Argument<string>("webpart-id") { Description = "WebPart ID" };
        var getCmd = new Command("get", "Get a webpart") { webpartArg, siteOption, pageIdOption, sectionIdReqOption, columnIdReqOption };
        getCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.GetWebPartAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionIdReqOption)!, parseResult.GetValue(columnIdReqOption)!, parseResult.GetValue(webpartArg)!);
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        // create
        var innerHtmlOption = new Option<string?>("--inner-html") { Description = "HTML content (creates a TextWebPart)" };
        var webPartTypeOption = new Option<string?>("--webpart-type") { Description = "Standard webpart type GUID (creates a StandardWebPart)" };
        var dataJsonOption = new Option<string?>("--data-json") { Description = "JSON data for standard webpart configuration" };
        var createCmd = new Command("create", "Add a webpart to a column") { siteOption, pageIdOption, sectionIdReqOption, columnIdReqOption, innerHtmlOption, webPartTypeOption, dataJsonOption };
        createCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.CreateWebPartAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionIdReqOption)!, parseResult.GetValue(columnIdReqOption)!, parseResult.GetValue(innerHtmlOption), parseResult.GetValue(webPartTypeOption), parseResult.GetValue(dataJsonOption));
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        // update
        var updateInnerHtmlOption = new Option<string?>("--inner-html") { Description = "New HTML content" };
        var updateDataJsonOption = new Option<string?>("--data-json") { Description = "New JSON data for standard webpart" };
        var updateCmd = new Command("update", "Update a webpart") { webpartArg, siteOption, pageIdOption, sectionIdReqOption, columnIdReqOption, updateInnerHtmlOption, updateDataJsonOption };
        updateCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.UpdateWebPartAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionIdReqOption)!, parseResult.GetValue(columnIdReqOption)!, parseResult.GetValue(webpartArg)!, parseResult.GetValue(updateInnerHtmlOption), parseResult.GetValue(updateDataJsonOption));
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        // delete
        var deleteCmd = new Command("delete", "Delete a webpart") { webpartArg, siteOption, pageIdOption, sectionIdReqOption, columnIdReqOption };
        deleteCmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await PageService.DeleteWebPartAsync(parseResult.GetValue(siteOption)!, parseResult.GetValue(pageIdOption)!, parseResult.GetValue(sectionIdReqOption)!, parseResult.GetValue(columnIdReqOption)!, parseResult.GetValue(webpartArg)!);
                OutputService.Print(result, format);
            }
            catch (ODataError ex) { OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message); Environment.ExitCode = 1; }
        });

        group.Subcommands.Add(listCmd);
        group.Subcommands.Add(getCmd);
        group.Subcommands.Add(createCmd);
        group.Subcommands.Add(updateCmd);
        group.Subcommands.Add(deleteCmd);
        return group;
    }
}
