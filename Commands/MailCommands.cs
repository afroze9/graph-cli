using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;


public static class MailCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var mailCommand = new Command("mail", "Email operations");

        mailCommand.Subcommands.Add(BuildList(formatOption));
        mailCommand.Subcommands.Add(BuildGet(formatOption));
        mailCommand.Subcommands.Add(BuildSearch(formatOption));
        mailCommand.Subcommands.Add(BuildSend(formatOption));
        mailCommand.Subcommands.Add(BuildDraft(formatOption));
        mailCommand.Subcommands.Add(BuildSendDraft(formatOption));
        mailCommand.Subcommands.Add(BuildMove(formatOption));
        mailCommand.Subcommands.Add(BuildDelete(formatOption));
        mailCommand.Subcommands.Add(BuildFolders(formatOption));

        return mailCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var folderOption = new Option<string?>("--folder") { Description = "Mail folder name (default: Inbox)" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 10, Description = "Number of messages to retrieve" };
        var cmd = new Command("list", "List messages") { folderOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var folder = parseResult.GetValue(folderOption);
            var top = parseResult.GetValue(topOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                MessageCollectionResponse? messages;

                if (!string.IsNullOrEmpty(folder))
                {
                    messages = await client.Me.MailFolders[folder].Messages.GetAsync(r =>
                    {
                        r.QueryParameters.Top = top;
                        r.QueryParameters.Select = ["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments"];
                        r.QueryParameters.Orderby = ["receivedDateTime desc"];
                    }, ct);
                }
                else
                {
                    messages = await client.Me.Messages.GetAsync(r =>
                    {
                        r.QueryParameters.Top = top;
                        r.QueryParameters.Select = ["id", "subject", "from", "receivedDateTime", "isRead", "hasAttachments"];
                        r.QueryParameters.Orderby = ["receivedDateTime desc"];
                    }, ct);
                }

                var results = messages?.Value?.Select(m => new
                {
                    m.Id,
                    m.Subject,
                    From = m.From?.EmailAddress?.Address,
                    m.ReceivedDateTime,
                    m.IsRead,
                    m.HasAttachments
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
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID" };
        var cmd = new Command("get", "Get message details") { messageIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var messageId = parseResult.GetValue(messageIdArg)!;
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var msg = await client.Me.Messages[messageId].GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id", "subject", "from", "toRecipients", "ccRecipients", "receivedDateTime", "body", "isRead", "hasAttachments", "importance"];
                }, ct);
                OutputService.Print(new
                {
                    msg!.Id,
                    msg.Subject,
                    From = msg.From?.EmailAddress?.Address,
                    To = msg.ToRecipients?.Select(r => r.EmailAddress?.Address).ToList(),
                    Cc = msg.CcRecipients?.Select(r => r.EmailAddress?.Address).ToList(),
                    msg.ReceivedDateTime,
                    BodyType = msg.Body?.ContentType?.ToString(),
                    Body = msg.Body?.Content,
                    msg.IsRead,
                    msg.HasAttachments,
                    Importance = msg.Importance?.ToString()
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

    private static Command BuildSearch(Option<string> formatOption)
    {
        var queryOption = new Option<string>("--query") { Description = "Search query", Required = true };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 10, Description = "Number of results" };
        var cmd = new Command("search", "Search messages") { queryOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var query = parseResult.GetValue(queryOption)!;
            var top = parseResult.GetValue(topOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var messages = await client.Me.Messages.GetAsync(r =>
                {
                    r.QueryParameters.Search = $"\"{query}\"";
                    r.QueryParameters.Top = top;
                    r.QueryParameters.Select = ["id", "subject", "from", "receivedDateTime", "isRead"];
                }, ct);
                var results = messages?.Value?.Select(m => new
                {
                    m.Id,
                    m.Subject,
                    From = m.From?.EmailAddress?.Address,
                    m.ReceivedDateTime,
                    m.IsRead
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

    private static Command BuildSend(Option<string> formatOption)
    {
        var toOption = new Option<string>("--to") { Description = "Comma-separated recipient emails", Required = true };
        var subjectOption = new Option<string>("--subject") { Description = "Email subject", Required = true };
        var bodyOption = new Option<string>("--body") { Description = "Email body", Required = true };
        var ccOption = new Option<string?>("--cc") { Description = "Comma-separated CC emails" };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Body content type: text or html" };
        var cmd = new Command("send", "Send an email") { toOption, subjectOption, bodyOption, ccOption, contentTypeOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var to = parseResult.GetValue(toOption)!;
            var subject = parseResult.GetValue(subjectOption)!;
            var body = parseResult.GetValue(bodyOption)!;
            var cc = parseResult.GetValue(ccOption);
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";

            var allRecipients = to.Split(',').Select(e => e.Trim()).ToList();
            if (!string.IsNullOrEmpty(cc))
                allRecipients.AddRange(cc.Split(',').Select(e => e.Trim()));

            if (!AllowedContactsService.CheckAllAndPrompt(allRecipients, "email"))
            {
                Environment.ExitCode = 1;
                return;
            }

            try
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

                await client.Me.SendMail.PostAsync(new SendMailPostRequestBody
                {
                    Message = message,
                    SaveToSentItems = true
                }, cancellationToken: ct);

                OutputService.Print(new { status = "sent", subject, to });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildDraft(Option<string> formatOption)
    {
        var toOption = new Option<string>("--to") { Description = "Comma-separated recipient emails", Required = true };
        var subjectOption = new Option<string>("--subject") { Description = "Email subject", Required = true };
        var bodyOption = new Option<string>("--body") { Description = "Email body", Required = true };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Body content type: text or html" };
        var cmd = new Command("draft", "Create a draft email") { toOption, subjectOption, bodyOption, contentTypeOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var to = parseResult.GetValue(toOption)!;
            var subject = parseResult.GetValue(subjectOption)!;
            var body = parseResult.GetValue(bodyOption)!;
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";

            var recipients = to.Split(',').Select(e => e.Trim());
            if (!AllowedContactsService.CheckAllAndPrompt(recipients, "email"))
            {
                Environment.ExitCode = 1;
                return;
            }

            try
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

                var draft = await client.Me.Messages.PostAsync(message, cancellationToken: ct);
                OutputService.Print(new { status = "draft_created", id = draft?.Id, subject });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildSendDraft(Option<string> formatOption)
    {
        var messageIdArg = new Argument<string>("message-id") { Description = "Draft message ID" };
        var cmd = new Command("send-draft", "Send an existing draft") { messageIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var messageId = parseResult.GetValue(messageIdArg)!;
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                await client.Me.Messages[messageId].Send.PostAsync(cancellationToken: ct);
                OutputService.Print(new { status = "sent", messageId });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildMove(Option<string> formatOption)
    {
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID" };
        var folderOption = new Option<string>("--folder") { Description = "Destination folder ID or well-known name", Required = true };
        var cmd = new Command("move", "Move message to folder") { messageIdArg, folderOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var messageId = parseResult.GetValue(messageIdArg)!;
            var folder = parseResult.GetValue(folderOption)!;
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var moved = await client.Me.Messages[messageId].Move.PostAsync(
                    new Microsoft.Graph.Me.Messages.Item.Move.MovePostRequestBody
                    {
                        DestinationId = folder
                    }, cancellationToken: ct);
                OutputService.Print(new { status = "moved", messageId = moved?.Id, folder });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildDelete(Option<string> formatOption)
    {
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID" };
        var cmd = new Command("delete", "Delete a message") { messageIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var messageId = parseResult.GetValue(messageIdArg)!;
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                await client.Me.Messages[messageId].DeleteAsync(cancellationToken: ct);
                OutputService.Print(new { status = "deleted", messageId });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildFolders(Option<string> formatOption)
    {
        var cmd = new Command("folders", "List mail folders");
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var folders = await client.Me.MailFolders.GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id", "displayName", "totalItemCount", "unreadItemCount"];
                }, ct);
                var results = folders?.Value?.Select(f => new
                {
                    f.Id,
                    f.DisplayName,
                    f.TotalItemCount,
                    f.UnreadItemCount
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
}
