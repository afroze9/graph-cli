using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class FilesCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var filesCommand = new Command("files", "OneDrive and SharePoint file operations");

        filesCommand.Subcommands.Add(BuildList(formatOption));
        filesCommand.Subcommands.Add(BuildGet(formatOption));
        filesCommand.Subcommands.Add(BuildDownload());
        filesCommand.Subcommands.Add(BuildSearch(formatOption));

        return filesCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var folderOption = new Option<string?>("--folder") { Description = "Folder item ID (default: root)" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname (e.g. contoso.sharepoint.com:/sites/team)" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of items to retrieve" };
        var cmd = new Command("list", "List files and folders in a drive") { folderOption, driveIdOption, siteOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var folder = parseResult.GetValue(folderOption);
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);
            var top = parseResult.GetValue(topOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var resolvedDrive = await ResolveDriveIdAsync(client, site, driveId, ct);
                var itemId = folder ?? "root";

                var items = await client.Drives[resolvedDrive].Items[itemId].Children.GetAsync(r =>
                {
                    r.QueryParameters.Top = top;
                    r.QueryParameters.Select = ["id", "name", "size", "lastModifiedDateTime", "folder", "file", "webUrl"];
                    r.QueryParameters.Orderby = ["name"];
                }, ct);

                var results = items?.Value?.Select(i => new
                {
                    i.Id,
                    i.Name,
                    Type = i.Folder != null ? "folder" : "file",
                    i.Size,
                    i.LastModifiedDateTime,
                    ChildCount = i.Folder?.ChildCount,
                    MimeType = i.File?.MimeType,
                    i.WebUrl
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
        var itemArg = new Argument<string>("item") { Description = "Item ID, or a sharing URL (https://...sharepoint.com/...)" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname" };
        var cmd = new Command("get", "Get file or folder metadata") { itemArg, driveIdOption, siteOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var item = parseResult.GetValue(itemArg)!;
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                DriveItem? driveItem;

                if (IsSharingUrl(item))
                {
                    var encoded = EncodeSharingUrl(item);
                    driveItem = await client.Shares[encoded].DriveItem.GetAsync(cancellationToken: ct);
                }
                else
                {
                    var resolvedDrive = await ResolveDriveIdAsync(client, site, driveId, ct);
                    driveItem = await client.Drives[resolvedDrive].Items[item].GetAsync(cancellationToken: ct);
                }

                if (driveItem == null)
                {
                    OutputService.PrintError("not_found", "Item not found.");
                    Environment.ExitCode = 1;
                    return;
                }

                OutputService.Print(new
                {
                    driveItem.Id,
                    driveItem.Name,
                    Type = driveItem.Folder != null ? "folder" : "file",
                    driveItem.Size,
                    driveItem.LastModifiedDateTime,
                    driveItem.CreatedDateTime,
                    ChildCount = driveItem.Folder?.ChildCount,
                    MimeType = driveItem.File?.MimeType,
                    driveItem.WebUrl,
                    CreatedBy = driveItem.CreatedBy?.User?.DisplayName,
                    LastModifiedBy = driveItem.LastModifiedBy?.User?.DisplayName,
                    DriveId = driveItem.ParentReference?.DriveId
                }, format);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildDownload()
    {
        var itemArg = new Argument<string>("item") { Description = "Item ID, or a sharing URL (https://...sharepoint.com/...)" };
        var outOption = new Option<string?>("--out") { Description = "Output file path (default: original filename in current directory)" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname" };
        var cmd = new Command("download", "Download a file") { itemArg, outOption, driveIdOption, siteOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var item = parseResult.GetValue(itemArg)!;
            var outPath = parseResult.GetValue(outOption);
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();

                string resolvedDrive;
                string resolvedItemId;

                if (IsSharingUrl(item))
                {
                    // Resolve the sharing URL to a drive item to get driveId + itemId
                    var encoded = EncodeSharingUrl(item);
                    var sharedItem = await client.Shares[encoded].DriveItem.GetAsync(r =>
                    {
                        r.QueryParameters.Select = ["id", "name", "size", "folder", "parentReference"];
                    }, ct);

                    if (sharedItem == null)
                    {
                        OutputService.PrintError("not_found", "Item not found.");
                        Environment.ExitCode = 1;
                        return;
                    }

                    if (sharedItem.Folder != null)
                    {
                        OutputService.PrintError("invalid_operation", "Cannot download a folder. Use 'files list' to see its contents.");
                        Environment.ExitCode = 1;
                        return;
                    }

                    resolvedDrive = sharedItem.ParentReference?.DriveId
                        ?? throw new InvalidOperationException("Could not determine drive ID from sharing URL.");
                    resolvedItemId = sharedItem.Id!;

                    var filePath = outPath ?? sharedItem.Name ?? "download";
                    var content = await client.Drives[resolvedDrive].Items[resolvedItemId].Content.GetAsync(cancellationToken: ct);
                    await WriteStreamToFileAsync(content, filePath, sharedItem.Size, ct);
                }
                else
                {
                    resolvedDrive = await ResolveDriveIdAsync(client, site, driveId, ct);
                    resolvedItemId = item;

                    // Get metadata to check if it's a folder and get the filename
                    var driveItem = await client.Drives[resolvedDrive].Items[resolvedItemId].GetAsync(r =>
                    {
                        r.QueryParameters.Select = ["id", "name", "size", "folder"];
                    }, ct);

                    if (driveItem == null)
                    {
                        OutputService.PrintError("not_found", "Item not found.");
                        Environment.ExitCode = 1;
                        return;
                    }

                    if (driveItem.Folder != null)
                    {
                        OutputService.PrintError("invalid_operation", "Cannot download a folder. Use 'files list' to see its contents.");
                        Environment.ExitCode = 1;
                        return;
                    }

                    var filePath = outPath ?? driveItem.Name ?? "download";
                    var content = await client.Drives[resolvedDrive].Items[resolvedItemId].Content.GetAsync(cancellationToken: ct);
                    await WriteStreamToFileAsync(content, filePath, driveItem.Size, ct);
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

    private static Command BuildSearch(Option<string> formatOption)
    {
        var queryArg = new Argument<string>("query") { Description = "Search query text" };
        var driveIdOption = new Option<string?>("--drive-id") { Description = "Drive ID (default: current user's OneDrive)" };
        var siteOption = new Option<string?>("--site") { Description = "SharePoint site ID or hostname" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of results to retrieve" };
        var cmd = new Command("search", "Search for files across OneDrive or SharePoint") { queryArg, driveIdOption, siteOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var query = parseResult.GetValue(queryArg)!;
            var driveId = parseResult.GetValue(driveIdOption);
            var site = parseResult.GetValue(siteOption);
            var top = parseResult.GetValue(topOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var resolvedDrive = await ResolveDriveIdAsync(client, site, driveId, ct);

                var results = await client.Drives[resolvedDrive].SearchWithQ(query).GetAsSearchWithQGetResponseAsync(r =>
                {
                    r.QueryParameters.Top = top;
                    r.QueryParameters.Select = ["id", "name", "size", "lastModifiedDateTime", "webUrl", "parentReference", "file", "folder"];
                }, ct);

                var items = results?.Value?.Select(i => new
                {
                    i.Id,
                    i.Name,
                    Type = i.Folder != null ? "folder" : "file",
                    i.Size,
                    i.LastModifiedDateTime,
                    MimeType = i.File?.MimeType,
                    Path = i.ParentReference?.Path,
                    i.WebUrl
                }).ToList();
                OutputService.Print(items, format);
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
    /// Resolves the drive ID from --site/--drive-id flags, or falls back to the user's OneDrive.
    /// </summary>
    private static async Task<string> ResolveDriveIdAsync(
        GraphServiceClient client, string? site, string? driveId, CancellationToken ct)
    {
        if (!string.IsNullOrEmpty(driveId))
            return driveId;

        if (!string.IsNullOrEmpty(site))
        {
            var siteObj = await client.Sites[site].GetAsync(r =>
            {
                r.QueryParameters.Select = ["id"];
            }, ct);

            if (siteObj?.Id == null)
                throw new InvalidOperationException($"Could not resolve site: {site}");

            var drive = await client.Sites[siteObj.Id].Drive.GetAsync(r =>
            {
                r.QueryParameters.Select = ["id"];
            }, ct);

            return drive?.Id ?? throw new InvalidOperationException($"No default drive found for site: {site}");
        }

        // Default: current user's OneDrive
        var myDrive = await client.Me.Drive.GetAsync(r =>
        {
            r.QueryParameters.Select = ["id"];
        }, ct);

        return myDrive?.Id ?? throw new InvalidOperationException("Could not resolve user's OneDrive.");
    }

    /// <summary>
    /// Checks if the input looks like a sharing URL (http/https).
    /// </summary>
    private static bool IsSharingUrl(string value) =>
        value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        value.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Encodes a sharing URL into the format expected by /shares/{encodedUrl}.
    /// See: https://learn.microsoft.com/en-us/graph/api/shares-get
    /// </summary>
    private static string EncodeSharingUrl(string url)
    {
        var base64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(url));
        // Convert to base64url and prepend "u!"
        var encoded = "u!" + base64.TrimEnd('=').Replace('/', '_').Replace('+', '-');
        return encoded;
    }

    private static async Task WriteStreamToFileAsync(Stream? content, string filePath, long? size, CancellationToken ct)
    {
        if (content == null)
        {
            OutputService.PrintError("no_content", "File has no downloadable content.");
            Environment.ExitCode = 1;
            return;
        }

        await using var fileStream = File.Create(filePath);
        await content.CopyToAsync(fileStream, ct);
        OutputService.Print(new { status = "downloaded", file = filePath, size });
    }
}
