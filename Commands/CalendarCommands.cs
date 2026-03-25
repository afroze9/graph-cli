using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class CalendarCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var calendarCommand = new Command("calendar", "Calendar operations");

        calendarCommand.Subcommands.Add(BuildList(formatOption));
        calendarCommand.Subcommands.Add(BuildEvents(formatOption));
        calendarCommand.Subcommands.Add(BuildCreateEvent(formatOption));
        calendarCommand.Subcommands.Add(BuildUpdateEvent(formatOption));
        calendarCommand.Subcommands.Add(BuildDeleteEvent(formatOption));
        calendarCommand.Subcommands.Add(BuildRespond(formatOption));

        return calendarCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var cmd = new Command("list", "List all calendars");
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var calendars = await client.Me.Calendars.GetAsync(r =>
                {
                    r.QueryParameters.Select = ["id", "name", "color", "isDefaultCalendar", "canEdit", "owner"];
                }, ct);
                var results = calendars?.Value?.Select(c => new
                {
                    c.Id,
                    c.Name,
                    Color = c.Color?.ToString(),
                    c.IsDefaultCalendar,
                    c.CanEdit,
                    OwnerName = c.Owner?.Name,
                    OwnerEmail = c.Owner?.Address
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

    private static Command BuildEvents(Option<string> formatOption)
    {
        var startOption = new Option<string?>("--start") { Description = "Start date (ISO 8601, default: today)" };
        var endOption = new Option<string?>("--end") { Description = "End date (ISO 8601, default: +7 days)" };
        var calendarIdOption = new Option<string?>("--calendar-id") { Description = "Specific calendar ID" };
        var topOption = new Option<int>("--top") { DefaultValueFactory = _ => 25, Description = "Number of events" };
        var cmd = new Command("events", "List calendar events") { startOption, endOption, calendarIdOption, topOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var start = parseResult.GetValue(startOption) ?? DateTime.Today.ToString("o");
            var end = parseResult.GetValue(endOption) ?? DateTime.Today.AddDays(7).ToString("o");
            var calendarId = parseResult.GetValue(calendarIdOption);
            var top = parseResult.GetValue(topOption);

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                EventCollectionResponse? events;

                if (!string.IsNullOrEmpty(calendarId))
                {
                    events = await client.Me.Calendars[calendarId].CalendarView.GetAsync(r =>
                    {
                        r.QueryParameters.StartDateTime = start;
                        r.QueryParameters.EndDateTime = end;
                        r.QueryParameters.Top = top;
                        r.QueryParameters.Select = ["id", "subject", "start", "end", "location", "organizer", "isAllDay", "isCancelled", "responseStatus"];
                        r.QueryParameters.Orderby = ["start/dateTime"];
                    }, ct);
                }
                else
                {
                    events = await client.Me.CalendarView.GetAsync(r =>
                    {
                        r.QueryParameters.StartDateTime = start;
                        r.QueryParameters.EndDateTime = end;
                        r.QueryParameters.Top = top;
                        r.QueryParameters.Select = ["id", "subject", "start", "end", "location", "organizer", "isAllDay", "isCancelled", "responseStatus"];
                        r.QueryParameters.Orderby = ["start/dateTime"];
                    }, ct);
                }

                var results = events?.Value?.Select(e => new
                {
                    e.Id,
                    e.Subject,
                    StartDateTime = e.Start?.DateTime,
                    StartTimeZone = e.Start?.TimeZone,
                    EndDateTime = e.End?.DateTime,
                    EndTimeZone = e.End?.TimeZone,
                    Location = e.Location?.DisplayName,
                    Organizer = e.Organizer?.EmailAddress?.Address,
                    e.IsAllDay,
                    e.IsCancelled,
                    Response = e.ResponseStatus?.Response?.ToString()
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

    private static Command BuildCreateEvent(Option<string> formatOption)
    {
        var subjectOption = new Option<string>("--subject") { Description = "Event subject", Required = true };
        var startOption = new Option<string>("--start") { Description = "Start datetime (ISO 8601)", Required = true };
        var endOption = new Option<string>("--end") { Description = "End datetime (ISO 8601)", Required = true };
        var attendeesOption = new Option<string?>("--attendees") { Description = "Comma-separated attendee emails" };
        var bodyOption = new Option<string?>("--body") { Description = "Event body/description" };
        var calendarIdOption = new Option<string?>("--calendar-id") { Description = "Calendar ID (default: primary)" };
        var cmd = new Command("create-event", "Create a calendar event")
            { subjectOption, startOption, endOption, attendeesOption, bodyOption, calendarIdOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var subject = parseResult.GetValue(subjectOption)!;
            var start = parseResult.GetValue(startOption)!;
            var end = parseResult.GetValue(endOption)!;
            var attendees = parseResult.GetValue(attendeesOption);
            var body = parseResult.GetValue(bodyOption);
            var calendarId = parseResult.GetValue(calendarIdOption);

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var newEvent = new Event
                {
                    Subject = subject,
                    Start = new DateTimeTimeZone { DateTime = start, TimeZone = TimeZoneInfo.Local.Id },
                    End = new DateTimeTimeZone { DateTime = end, TimeZone = TimeZoneInfo.Local.Id }
                };

                if (!string.IsNullOrEmpty(body))
                    newEvent.Body = new ItemBody { ContentType = BodyType.Text, Content = body };

                if (!string.IsNullOrEmpty(attendees))
                {
                    newEvent.Attendees = attendees.Split(',').Select(e => new Attendee
                    {
                        EmailAddress = new EmailAddress { Address = e.Trim() },
                        Type = AttendeeType.Required
                    }).ToList();
                }

                Event? created;
                if (!string.IsNullOrEmpty(calendarId))
                    created = await client.Me.Calendars[calendarId].Events.PostAsync(newEvent, cancellationToken: ct);
                else
                    created = await client.Me.Events.PostAsync(newEvent, cancellationToken: ct);

                OutputService.Print(new { status = "created", id = created?.Id, subject, start, end });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildUpdateEvent(Option<string> formatOption)
    {
        var eventIdArg = new Argument<string>("event-id") { Description = "Event ID" };
        var subjectOption = new Option<string?>("--subject") { Description = "New subject" };
        var startOption = new Option<string?>("--start") { Description = "New start datetime" };
        var endOption = new Option<string?>("--end") { Description = "New end datetime" };
        var bodyOption = new Option<string?>("--body") { Description = "New body" };
        var cmd = new Command("update-event", "Update a calendar event") { eventIdArg, subjectOption, startOption, endOption, bodyOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var eventId = parseResult.GetValue(eventIdArg)!;
            var subject = parseResult.GetValue(subjectOption);
            var start = parseResult.GetValue(startOption);
            var end = parseResult.GetValue(endOption);
            var body = parseResult.GetValue(bodyOption);

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var update = new Event();

                if (subject != null) update.Subject = subject;
                if (start != null) update.Start = new DateTimeTimeZone { DateTime = start, TimeZone = TimeZoneInfo.Local.Id };
                if (end != null) update.End = new DateTimeTimeZone { DateTime = end, TimeZone = TimeZoneInfo.Local.Id };
                if (body != null) update.Body = new ItemBody { ContentType = BodyType.Text, Content = body };

                var updated = await client.Me.Events[eventId].PatchAsync(update, cancellationToken: ct);
                OutputService.Print(new { status = "updated", id = updated?.Id });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildDeleteEvent(Option<string> formatOption)
    {
        var eventIdArg = new Argument<string>("event-id") { Description = "Event ID" };
        var cmd = new Command("delete-event", "Delete a calendar event") { eventIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var eventId = parseResult.GetValue(eventIdArg)!;
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                await client.Me.Events[eventId].DeleteAsync(cancellationToken: ct);
                OutputService.Print(new { status = "deleted", eventId });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildRespond(Option<string> formatOption)
    {
        var eventIdArg = new Argument<string>("event-id") { Description = "Event ID" };
        var actionOption = new Option<string>("--action") { Description = "Response: accept, decline, or tentative", Required = true };
        var commentOption = new Option<string?>("--comment") { Description = "Optional response comment" };
        var cmd = new Command("respond", "Respond to a meeting invitation") { eventIdArg, actionOption, commentOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var eventId = parseResult.GetValue(eventIdArg)!;
            var action = parseResult.GetValue(actionOption)!.ToLower();
            var comment = parseResult.GetValue(commentOption);

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                switch (action)
                {
                    case "accept":
                        await client.Me.Events[eventId].Accept.PostAsync(
                            new Microsoft.Graph.Me.Events.Item.Accept.AcceptPostRequestBody
                            {
                                Comment = comment,
                                SendResponse = true
                            }, cancellationToken: ct);
                        break;
                    case "decline":
                        await client.Me.Events[eventId].Decline.PostAsync(
                            new Microsoft.Graph.Me.Events.Item.Decline.DeclinePostRequestBody
                            {
                                Comment = comment,
                                SendResponse = true
                            }, cancellationToken: ct);
                        break;
                    case "tentative":
                        await client.Me.Events[eventId].TentativelyAccept.PostAsync(
                            new Microsoft.Graph.Me.Events.Item.TentativelyAccept.TentativelyAcceptPostRequestBody
                            {
                                Comment = comment,
                                SendResponse = true
                            }, cancellationToken: ct);
                        break;
                    default:
                        OutputService.PrintError("invalid_action", "Action must be: accept, decline, or tentative");
                        Environment.ExitCode = 1;
                        return;
                }
                OutputService.Print(new { status = "responded", eventId, action });
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
