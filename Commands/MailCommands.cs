using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class MailCommands
{
    public static Command Build(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var mailCommand = new Command("mail", "Email operations");

        mailCommand.Subcommands.Add(BuildList(formatOption, timezoneOption));
        mailCommand.Subcommands.Add(BuildGet(formatOption, timezoneOption));
        mailCommand.Subcommands.Add(BuildSearch(formatOption, timezoneOption));
        mailCommand.Subcommands.Add(BuildSend(formatOption));
        mailCommand.Subcommands.Add(BuildDraft(formatOption));
        mailCommand.Subcommands.Add(BuildSendDraft(formatOption));
        mailCommand.Subcommands.Add(BuildMove(formatOption));
        mailCommand.Subcommands.Add(BuildDelete(formatOption));
        mailCommand.Subcommands.Add(BuildFolders(formatOption));
        mailCommand.Subcommands.Add(BuildMarkRead(formatOption));
        mailCommand.Subcommands.Add(BuildAttachments(formatOption));
        mailCommand.Subcommands.Add(BuildDownloadAttachment());

        return mailCommand;
    }

    private static Command BuildList(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var folderOption = new Option<string?>("--folder") { Description = "Mail folder name (default: Inbox)" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 10, Description = "Number of messages to retrieve" };
        var cmd = new Command("list", "List messages") { folderOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var folder = parseResult.GetValue(folderOption);
            var top = parseResult.GetValue(topOption);
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));
            try
            {
                var result = await MailService.ListAsync(folder, top, tz);
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

    private static Command BuildGet(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID" };
        var cmd = new Command("get", "Get message details") { messageIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var messageId = parseResult.GetValue(messageIdArg)!;
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));
            try
            {
                var result = await MailService.GetAsync(messageId, tz);
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

    private static Command BuildSearch(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var queryOption = new Option<string>("--query") { Description = "Search query", Required = true };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 10, Description = "Number of results" };
        var cmd = new Command("search", "Search messages") { queryOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var query = parseResult.GetValue(queryOption)!;
            var top = parseResult.GetValue(topOption);
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));
            try
            {
                var result = await MailService.SearchAsync(query, top, tz);
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
            var cc = parseResult.GetValue(ccOption);

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
                var result = await MailService.SendAsync(
                    to,
                    parseResult.GetValue(subjectOption)!,
                    parseResult.GetValue(bodyOption)!,
                    cc,
                    parseResult.GetValue(contentTypeOption) ?? "text");
                OutputService.Print(result);
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
            var recipients = to.Split(',').Select(e => e.Trim());
            if (!AllowedContactsService.CheckAllAndPrompt(recipients, "email"))
            {
                Environment.ExitCode = 1;
                return;
            }

            try
            {
                var result = await MailService.DraftAsync(
                    to,
                    parseResult.GetValue(subjectOption)!,
                    parseResult.GetValue(bodyOption)!,
                    parseResult.GetValue(contentTypeOption) ?? "text");
                OutputService.Print(result);
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
            try
            {
                var result = await MailService.SendDraftAsync(parseResult.GetValue(messageIdArg)!);
                OutputService.Print(result);
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
        var messageIdsArg = new Argument<string[]>("message-id") { Description = "One or more message IDs", Arity = ArgumentArity.OneOrMore };
        var folderOption = new Option<string>("--folder") { Description = "Destination folder ID or well-known name", Required = true };
        var cmd = new Command("move", "Move one or more messages to a folder") { messageIdsArg, folderOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var result = await MailService.MoveAsync(
                    parseResult.GetValue(messageIdsArg)!,
                    parseResult.GetValue(folderOption)!);
                OutputService.Print(result);
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
        var messageIdsArg = new Argument<string[]>("message-id") { Description = "One or more message IDs", Arity = ArgumentArity.OneOrMore };
        var cmd = new Command("delete", "Delete one or more messages") { messageIdsArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var result = await MailService.DeleteAsync(parseResult.GetValue(messageIdsArg)!);
                OutputService.Print(result);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildMarkRead(Option<string> formatOption)
    {
        var messageIdsArg = new Argument<string[]>("message-id") { Description = "One or more message IDs", Arity = ArgumentArity.OneOrMore };
        var unreadOption = new Option<bool>("--unread") { Description = "Mark as unread instead of read" };
        var cmd = new Command("mark-read", "Mark one or more messages as read or unread") { messageIdsArg, unreadOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var result = await MailService.MarkReadAsync(
                    parseResult.GetValue(messageIdsArg)!,
                    parseResult.GetValue(unreadOption));
                OutputService.Print(result);
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
        var parentOption = new Option<string?>("--parent") { Description = "Parent folder ID or well-known name to list child folders" };
        var cmd = new Command("folders", "List mail folders") { parentOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await MailService.FoldersAsync(parseResult.GetValue(parentOption));
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

    private static Command BuildAttachments(Option<string> formatOption)
    {
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID" };
        var cmd = new Command("attachments", "List attachments on a message") { messageIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var result = await MailService.AttachmentsAsync(parseResult.GetValue(messageIdArg)!);
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

    private static Command BuildDownloadAttachment()
    {
        var messageIdArg = new Argument<string>("message-id") { Description = "Message ID" };
        var attachmentIdArg = new Argument<string>("attachment-id") { Description = "Attachment ID" };
        var outOption = new Option<string?>("--out") { Description = "Output file path (default: attachment name in current directory)" };
        var cmd = new Command("download-attachment", "Download an attachment to a file") { messageIdArg, attachmentIdArg, outOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var result = await MailService.DownloadAttachmentAsync(
                    parseResult.GetValue(messageIdArg)!,
                    parseResult.GetValue(attachmentIdArg)!,
                    parseResult.GetValue(outOption));
                OutputService.Print(result);
            }
            catch (InvalidOperationException ex)
            {
                OutputService.PrintError("unsupported", ex.Message);
                Environment.ExitCode = 1;
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
