using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace GraphCli.Services;

public static class FileService
{
    public static async Task<object> ListAsync(string? folder, string? driveId, string? site, int top)
    {
        var client = await GraphClientProvider.CreateAsync();
        var resolvedDrive = await ResolveDriveIdAsync(client, site, driveId);
        var itemId = folder ?? "root";

        var items = await client.Drives[resolvedDrive].Items[itemId].Children.GetAsync(r =>
        {
            r.QueryParameters.Top = top;
            r.QueryParameters.Select = ["id", "name", "size", "lastModifiedDateTime", "folder", "file", "webUrl"];
            r.QueryParameters.Orderby = ["name"];
        });

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

        if (results != null)
        {
            FileCacheService.UpsertMany(results.Select(r =>
                (r.Id!, r.Name!, r.Type, r.Size, r.MimeType, r.WebUrl)));
        }

        return results ?? [];
    }

    public static async Task<object> GetAsync(string item, string? driveId, string? site)
    {
        var client = await GraphClientProvider.CreateAsync();
        DriveItem? driveItem;

        if (IsSharingUrl(item))
        {
            var encoded = EncodeSharingUrl(item);
            driveItem = await client.Shares[encoded].DriveItem.GetAsync();
        }
        else
        {
            var resolvedDrive = await ResolveDriveIdAsync(client, site, driveId);
            driveItem = await client.Drives[resolvedDrive].Items[item].GetAsync();
        }

        if (driveItem == null)
            throw new InvalidOperationException("Item not found.");

        return new
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
        };
    }

    public static async Task<object> SearchAsync(string query, string? driveId, string? site, int top, bool refresh)
    {
        // Check cache first (only when no site/drive filters and not refreshing)
        if (!refresh && string.IsNullOrEmpty(site) && string.IsNullOrEmpty(driveId))
        {
            var cached = FileCacheService.Search(query, top);
            if (cached.Count > 0)
            {
                return cached.Select(f => new
                {
                    f.Id,
                    f.Name,
                    f.Type,
                    f.Size,
                    f.MimeType,
                    f.WebUrl
                }).ToList();
            }
        }

        var client = await GraphClientProvider.CreateAsync();
        var resolvedDrive = await ResolveDriveIdAsync(client, site, driveId);

        var results = await client.Drives[resolvedDrive].SearchWithQ(query).GetAsSearchWithQGetResponseAsync(r =>
        {
            r.QueryParameters.Top = top;
            r.QueryParameters.Select = ["id", "name", "size", "lastModifiedDateTime", "webUrl", "parentReference", "file", "folder"];
        });

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

        if (items != null)
        {
            FileCacheService.UpsertMany(items.Select(i =>
                (i.Id!, i.Name!, i.Type, i.Size, i.MimeType, i.WebUrl)));
        }

        return items ?? [];
    }

    public static async Task<object> ShareAsync(string item, string recipients, string role, string? message, string? driveId, string? site)
    {
        var client = await GraphClientProvider.CreateAsync();

        string resolvedDrive;
        string resolvedItemId;

        if (IsSharingUrl(item))
        {
            var encoded = EncodeSharingUrl(item);
            var sharedItem = await client.Shares[encoded].DriveItem.GetAsync(r =>
            {
                r.QueryParameters.Select = ["id", "name", "parentReference"];
            });

            if (sharedItem == null)
                throw new InvalidOperationException("Item not found.");

            resolvedDrive = sharedItem.ParentReference?.DriveId
                ?? throw new InvalidOperationException("Could not determine drive ID from sharing URL.");
            resolvedItemId = sharedItem.Id!;
        }
        else
        {
            resolvedDrive = await ResolveDriveIdAsync(client, site, driveId);
            resolvedItemId = item;
        }

        var emails = recipients.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var inviteRecipients = emails.Select(email => new DriveRecipient
        {
            Email = email
        }).ToList();

        var invite = new Microsoft.Graph.Drives.Item.Items.Item.Invite.InvitePostRequestBody
        {
            Recipients = inviteRecipients,
            Roles = [role],
            RequireSignIn = true,
            SendInvitation = false,
            Message = message
        };

        var permissions = await client.Drives[resolvedDrive].Items[resolvedItemId]
            .Invite.PostAsInvitePostResponseAsync(invite);

        var permResults = permissions?.Value?.Select(p => new
        {
            p.Id,
            Role = p.Roles?.FirstOrDefault(),
            Email = p.GrantedToV2?.User?.DisplayName
                ?? p.Invitation?.Email,
            Link = p.Link?.WebUrl
        }).ToList();

        return new { status = "shared", item = resolvedItemId, permissions = permResults };
    }

    public static async Task<object> DownloadAsync(string item, string? outPath, string? driveId, string? site)
    {
        var client = await GraphClientProvider.CreateAsync();

        string resolvedDrive;
        string resolvedItemId;

        if (IsSharingUrl(item))
        {
            var encoded = EncodeSharingUrl(item);
            var sharedItem = await client.Shares[encoded].DriveItem.GetAsync(r =>
            {
                r.QueryParameters.Select = ["id", "name", "size", "folder", "parentReference"];
            });

            if (sharedItem == null)
                throw new InvalidOperationException("Item not found.");

            if (sharedItem.Folder != null)
                throw new InvalidOperationException("Cannot download a folder. Use 'files list' to see its contents.");

            resolvedDrive = sharedItem.ParentReference?.DriveId
                ?? throw new InvalidOperationException("Could not determine drive ID from sharing URL.");
            resolvedItemId = sharedItem.Id!;

            var filePath = outPath ?? Path.GetFileName(sharedItem.Name) ?? "download";
            var content = await client.Drives[resolvedDrive].Items[resolvedItemId].Content.GetAsync();
            await WriteStreamToFileAsync(content, filePath, sharedItem.Size);
            return new { status = "downloaded", file = filePath, size = sharedItem.Size };
        }
        else
        {
            resolvedDrive = await ResolveDriveIdAsync(client, site, driveId);
            resolvedItemId = item;

            var driveItem = await client.Drives[resolvedDrive].Items[resolvedItemId].GetAsync(r =>
            {
                r.QueryParameters.Select = ["id", "name", "size", "folder"];
            });

            if (driveItem == null)
                throw new InvalidOperationException("Item not found.");

            if (driveItem.Folder != null)
                throw new InvalidOperationException("Cannot download a folder. Use 'files list' to see its contents.");

            var filePath = outPath ?? Path.GetFileName(driveItem.Name) ?? "download";
            var content = await client.Drives[resolvedDrive].Items[resolvedItemId].Content.GetAsync();
            await WriteStreamToFileAsync(content, filePath, driveItem.Size);
            return new { status = "downloaded", file = filePath, size = driveItem.Size };
        }
    }

    private static async Task<string> ResolveDriveIdAsync(
        GraphServiceClient client, string? site, string? driveId, CancellationToken ct = default)
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

        var myDrive = await client.Me.Drive.GetAsync(r =>
        {
            r.QueryParameters.Select = ["id"];
        }, ct);

        return myDrive?.Id ?? throw new InvalidOperationException("Could not resolve user's OneDrive.");
    }

    private static bool IsSharingUrl(string value) =>
        value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        value.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

    private static string EncodeSharingUrl(string url)
    {
        var base64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(url));
        return "u!" + base64.TrimEnd('=').Replace('/', '_').Replace('+', '-');
    }

    private static async Task WriteStreamToFileAsync(Stream? content, string filePath, long? size)
    {
        if (content == null)
            throw new InvalidOperationException("File has no downloadable content.");

        await using var fileStream = File.Create(filePath);
        await content.CopyToAsync(fileStream);
    }
}
