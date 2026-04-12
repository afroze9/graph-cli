using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;

namespace GraphCli.Services;

public static class MailService
{
    public static async Task<object> ListAsync(string? folder, int top, string tz)
    {
        var client = await GraphClientProvider.CreateAsync();
        MessageCollectionResponse? messages;
        string[] select = ["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments"];

        if (!string.IsNullOrEmpty(folder))
        {
            messages = await client.Me.MailFolders[folder].Messages.GetAsync(r =>
            {
                r.QueryParameters.Top = top;
                r.QueryParameters.Select = select;
                r.QueryParameters.Orderby = ["receivedDateTime desc"];
            });
        }
        else
        {
            messages = await client.Me.Messages.GetAsync(r =>
            {
                r.QueryParameters.Top = top;
                r.QueryParameters.Select = select;
                r.QueryParameters.Orderby = ["receivedDateTime desc"];
            });
        }

        return messages?.Value?.Select(m => new
        {
            m.Id, m.Subject,
            From = m.From?.EmailAddress?.Address,
            ReceivedDateTime = TimeZoneService.ConvertToTimeZone(m.ReceivedDateTime, tz),
            m.IsRead, m.HasAttachments
        }).ToList() ?? [];
    }

    public static async Task<object> GetAsync(string messageId, string tz)
    {
        var client = await GraphClientProvider.CreateAsync();
        var msg = await client.Me.Messages[messageId].GetAsync(r =>
        {
            r.QueryParameters.Select = ["id", "subject", "from", "toRecipients", "ccRecipients", "receivedDateTime", "body", "isRead", "hasAttachments", "importance"];
        });
        return new
        {
            msg!.Id, msg.Subject,
            From = msg.From?.EmailAddress?.Address,
            To = msg.ToRecipients?.Select(r => r.EmailAddress?.Address).ToList(),
            Cc = msg.CcRecipients?.Select(r => r.EmailAddress?.Address).ToList(),
            ReceivedDateTime = TimeZoneService.ConvertToTimeZone(msg.ReceivedDateTime, tz),
            BodyType = msg.Body?.ContentType?.ToString(),
            Body = msg.Body?.Content,
            msg.IsRead, msg.HasAttachments,
            Importance = msg.Importance?.ToString()
        };
    }

    public static async Task<object> SearchAsync(string query, int top, string tz)
    {
        var client = await GraphClientProvider.CreateAsync();
        var messages = await client.Me.Messages.GetAsync(r =>
        {
            r.QueryParameters.Search = $"\"{query.Replace("\"", "\\\"")}\"";
            r.QueryParameters.Top = top;
            r.QueryParameters.Select = ["id", "subject", "from", "receivedDateTime", "isRead"];
        });
        return messages?.Value?.Select(m => new
        {
            m.Id, m.Subject,
            From = m.From?.EmailAddress?.Address,
            ReceivedDateTime = TimeZoneService.ConvertToTimeZone(m.ReceivedDateTime, tz),
            m.IsRead
        }).ToList() ?? [];
    }

    public static async Task<object> SendAsync(string to, string subject, string body, string? cc, string contentType, string[]? attachments = null)
    {
        var client = await GraphClientProvider.CreateAsync();
        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = contentType == "html" ? BodyType.Html : BodyType.Text,
                Content = body
            },
            ToRecipients = to.Split(',').Select(e => new Recipient
            {
                EmailAddress = new EmailAddress { Address = e.Trim() }
            }).ToList()
        };

        if (!string.IsNullOrEmpty(cc))
        {
            message.CcRecipients = cc.Split(',').Select(e => new Recipient
            {
                EmailAddress = new EmailAddress { Address = e.Trim() }
            }).ToList();
        }

        if (attachments is { Length: > 0 })
        {
            message.Attachments = attachments.Select(filePath =>
            {
                var fullPath = Path.GetFullPath(filePath);
                if (!File.Exists(fullPath))
                    throw new FileNotFoundException($"Attachment file not found: {fullPath}");

                var bytes = File.ReadAllBytes(fullPath);
                return (Attachment)new FileAttachment
                {
                    OdataType = "#microsoft.graph.fileAttachment",
                    Name = Path.GetFileName(fullPath),
                    ContentType = MimeTypeMap.GetMimeType(fullPath),
                    ContentBytes = bytes
                };
            }).ToList();
        }

        await client.Me.SendMail.PostAsync(new SendMailPostRequestBody
        {
            Message = message,
            SaveToSentItems = true
        });
        return new { status = "sent", subject, to };
    }

    private static class MimeTypeMap
    {
        private static readonly Dictionary<string, string> MimeTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            [".pdf"] = "application/pdf",
            [".doc"] = "application/msword",
            [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            [".xls"] = "application/vnd.ms-excel",
            [".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            [".ppt"] = "application/vnd.ms-powerpoint",
            [".pptx"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            [".txt"] = "text/plain",
            [".csv"] = "text/csv",
            [".html"] = "text/html",
            [".htm"] = "text/html",
            [".json"] = "application/json",
            [".xml"] = "application/xml",
            [".zip"] = "application/zip",
            [".png"] = "image/png",
            [".jpg"] = "image/jpeg",
            [".jpeg"] = "image/jpeg",
            [".gif"] = "image/gif",
            [".svg"] = "image/svg+xml",
            [".mp4"] = "video/mp4",
            [".mp3"] = "audio/mpeg",
        };

        public static string GetMimeType(string filePath)
        {
            var ext = Path.GetExtension(filePath);
            return MimeTypes.GetValueOrDefault(ext, "application/octet-stream");
        }
    }

    public static async Task<object> DraftAsync(string to, string subject, string body, string contentType)
    {
        var client = await GraphClientProvider.CreateAsync();
        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = contentType == "html" ? BodyType.Html : BodyType.Text,
                Content = body
            },
            ToRecipients = to.Split(',').Select(e => new Recipient
            {
                EmailAddress = new EmailAddress { Address = e.Trim() }
            }).ToList()
        };

        var draft = await client.Me.Messages.PostAsync(message);
        return new { status = "draft_created", id = draft?.Id, subject };
    }

    public static async Task<object> SendDraftAsync(string messageId)
    {
        var client = await GraphClientProvider.CreateAsync();
        await client.Me.Messages[messageId].Send.PostAsync();
        return new { status = "sent", messageId };
    }

    public static async Task<object> MoveAsync(string[] messageIds, string folder)
    {
        var client = await GraphClientProvider.CreateAsync();
        var results = new List<object>();
        foreach (var messageId in messageIds)
        {
            try
            {
                var moved = await client.Me.Messages[messageId].Move.PostAsync(
                    new Microsoft.Graph.Me.Messages.Item.Move.MovePostRequestBody
                    {
                        DestinationId = folder
                    });
                results.Add(new { status = "moved", messageId = moved?.Id, folder });
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                results.Add(new { status = "error", messageId, error = ex.Error?.Message ?? ex.Message });
            }
        }
        return messageIds.Length == 1 ? results[0] : results;
    }

    public static async Task<object> DeleteAsync(string[] messageIds)
    {
        var client = await GraphClientProvider.CreateAsync();
        var results = new List<object>();
        foreach (var messageId in messageIds)
        {
            try
            {
                await client.Me.Messages[messageId].DeleteAsync();
                results.Add(new { status = "deleted", messageId });
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                results.Add(new { status = "error", messageId, error = ex.Error?.Message ?? ex.Message });
            }
        }
        return messageIds.Length == 1 ? results[0] : results;
    }

    public static async Task<object> MarkReadAsync(string[] messageIds, bool unread)
    {
        var client = await GraphClientProvider.CreateAsync();
        var isRead = !unread;
        var results = new List<object>();
        foreach (var messageId in messageIds)
        {
            try
            {
                await client.Me.Messages[messageId].PatchAsync(new Message { IsRead = isRead });
                results.Add(new { status = isRead ? "marked_read" : "marked_unread", messageId });
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                results.Add(new { status = "error", messageId, error = ex.Error?.Message ?? ex.Message });
            }
        }
        return messageIds.Length == 1 ? results[0] : results;
    }

    public static async Task<object> FoldersAsync(string? parent)
    {
        var client = await GraphClientProvider.CreateAsync();
        MailFolderCollectionResponse? folders;

        if (!string.IsNullOrEmpty(parent))
        {
            folders = await client.Me.MailFolders[parent].ChildFolders.GetAsync(r =>
            {
                r.QueryParameters.Select = ["id", "displayName", "parentFolderId", "totalItemCount", "unreadItemCount", "childFolderCount"];
            });
        }
        else
        {
            folders = await client.Me.MailFolders.GetAsync(r =>
            {
                r.QueryParameters.Select = ["id", "displayName", "totalItemCount", "unreadItemCount", "childFolderCount"];
            });
        }

        return folders?.Value?.Select(f => new
        {
            f.Id, f.DisplayName, f.ParentFolderId,
            f.TotalItemCount, f.UnreadItemCount, f.ChildFolderCount
        }).ToList() ?? [];
    }

    public static async Task<object> AttachmentsAsync(string messageId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var attachments = await client.Me.Messages[messageId].Attachments.GetAsync(r =>
        {
            r.QueryParameters.Select = ["id", "name", "contentType", "size", "isInline"];
        });
        return attachments?.Value?.Select(a => new
        {
            a.Id, a.Name, a.ContentType, a.Size, a.IsInline
        }).ToList() ?? [];
    }

    public static async Task<object> DownloadAttachmentAsync(string messageId, string attachmentId, string? outPath)
    {
        var client = await GraphClientProvider.CreateAsync();
        var attachment = await client.Me.Messages[messageId].Attachments[attachmentId].GetAsync();

        if (attachment is FileAttachment fileAttachment && fileAttachment.ContentBytes != null)
        {
            var fileName = outPath ?? Path.GetFileName(fileAttachment.Name) ?? "attachment";
            await File.WriteAllBytesAsync(fileName, fileAttachment.ContentBytes);
            return new { status = "downloaded", file = fileName, size = fileAttachment.ContentBytes.Length };
        }

        throw new InvalidOperationException("Only file attachments can be downloaded. Item and reference attachments are not supported.");
    }
}
