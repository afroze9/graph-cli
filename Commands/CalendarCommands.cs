using System.CommandLine;
using System.Xml;
using GraphCli.Services;
using Microsoft.Graph;
using Microsoft.Graph.Me.Calendar.GetSchedule;
using Microsoft.Graph.Me.FindMeetingTimes;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class CalendarCommands
{
    public static Command Build(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var calendarCommand = new Command("calendar", "Calendar operations");

        calendarCommand.Subcommands.Add(BuildList(formatOption));
        calendarCommand.Subcommands.Add(BuildEvents(formatOption, timezoneOption));
        calendarCommand.Subcommands.Add(BuildGetEvent(formatOption, timezoneOption));
        calendarCommand.Subcommands.Add(BuildCreateEvent(formatOption, timezoneOption));
        calendarCommand.Subcommands.Add(BuildUpdateEvent(formatOption, timezoneOption));
        calendarCommand.Subcommands.Add(BuildDeleteEvent(formatOption));
        calendarCommand.Subcommands.Add(BuildRespond(formatOption));
        calendarCommand.Subcommands.Add(BuildFindTimes(formatOption, timezoneOption));
        calendarCommand.Subcommands.Add(BuildSchedule(formatOption, timezoneOption));

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

    private static Command BuildEvents(Option<string> formatOption, Option<string?> timezoneOption)
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
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));

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
                        r.QueryParameters.Select = ["id", "subject", "start", "end", "location", "organizer", "isAllDay", "isCancelled", "responseStatus", "categories"];
                        r.QueryParameters.Orderby = ["start/dateTime"];
                        r.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
                    }, ct);
                }
                else
                {
                    events = await client.Me.CalendarView.GetAsync(r =>
                    {
                        r.QueryParameters.StartDateTime = start;
                        r.QueryParameters.EndDateTime = end;
                        r.QueryParameters.Top = top;
                        r.QueryParameters.Select = ["id", "subject", "start", "end", "location", "organizer", "isAllDay", "isCancelled", "responseStatus", "categories"];
                        r.QueryParameters.Orderby = ["start/dateTime"];
                        r.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
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
                    Response = e.ResponseStatus?.Response?.ToString(),
                    Categories = e.Categories
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

    private static Command BuildGetEvent(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var eventIdArg = new Argument<string>("event-id") { Description = "Event ID" };
        var cmd = new Command("get-event", "Get full event details including attendees and body") { eventIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var eventId = parseResult.GetValue(eventIdArg)!;
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var e = await client.Me.Events[eventId].GetAsync(r =>
                {
                    r.QueryParameters.Select =
                    [
                        "id", "subject", "body", "bodyPreview", "start", "end",
                        "location", "locations", "organizer", "attendees",
                        "isOnlineMeeting", "onlineMeeting", "onlineMeetingProvider",
                        "importance", "sensitivity", "isAllDay", "isCancelled",
                        "responseStatus", "categories", "hasAttachments",
                        "recurrence", "webLink"
                    ];
                    r.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
                }, ct);

                OutputService.Print(new
                {
                    e!.Id,
                    e.Subject,
                    BodyType = e.Body?.ContentType?.ToString(),
                    Body = e.Body?.Content,
                    e.BodyPreview,
                    StartDateTime = e.Start?.DateTime,
                    StartTimeZone = e.Start?.TimeZone,
                    EndDateTime = e.End?.DateTime,
                    EndTimeZone = e.End?.TimeZone,
                    Location = e.Location?.DisplayName,
                    Locations = e.Locations?.Select(l => l.DisplayName).ToList(),
                    Organizer = e.Organizer?.EmailAddress?.Address,
                    Attendees = e.Attendees?.Select(a => new
                    {
                        Name = a.EmailAddress?.Name,
                        Email = a.EmailAddress?.Address,
                        Type = a.Type?.ToString(),
                        Response = a.Status?.Response?.ToString(),
                        ResponseTime = a.Status?.Time?.ToString("o")
                    }).ToList(),
                    e.IsOnlineMeeting,
                    JoinUrl = e.OnlineMeeting?.JoinUrl,
                    OnlineMeetingProvider = e.OnlineMeetingProvider?.ToString(),
                    Importance = e.Importance?.ToString(),
                    Sensitivity = e.Sensitivity?.ToString(),
                    e.IsAllDay,
                    e.IsCancelled,
                    Response = e.ResponseStatus?.Response?.ToString(),
                    e.Categories,
                    e.HasAttachments,
                    Recurrence = e.Recurrence != null ? new
                    {
                        Pattern = e.Recurrence.Pattern?.Type?.ToString(),
                        Interval = e.Recurrence.Pattern?.Interval,
                        DaysOfWeek = e.Recurrence.Pattern?.DaysOfWeek?.Select(d => d.ToString()).ToList(),
                        RangeType = e.Recurrence.Range?.Type?.ToString(),
                        RangeStart = e.Recurrence.Range?.StartDate?.ToString(),
                        RangeEnd = e.Recurrence.Range?.EndDate?.ToString()
                    } : null,
                    e.WebLink
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

    private static Command BuildCreateEvent(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var subjectOption = new Option<string>("--subject") { Description = "Event subject", Required = true };
        var startOption = new Option<string>("--start") { Description = "Start datetime (ISO 8601)", Required = true };
        var endOption = new Option<string>("--end") { Description = "End datetime (ISO 8601)", Required = true };
        var attendeesOption = new Option<string?>("--attendees") { Description = "Comma-separated attendee emails" };
        var bodyOption = new Option<string?>("--body") { Description = "Event body/description" };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Body content type: text or html" };
        var categoriesOption = new Option<string?>("--categories") { Description = "Comma-separated category names" };
        var locationOption = new Option<string?>("--location") { Description = "Event location (e.g. room name or address)" };
        var onlineMeetingOption = new Option<bool>("--online-meeting") { Description = "Generate a Teams online meeting link" };
        var calendarIdOption = new Option<string?>("--calendar-id") { Description = "Calendar ID (default: primary)" };
        var cmd = new Command("create-event", "Create a calendar event")
            { subjectOption, startOption, endOption, attendeesOption, bodyOption, contentTypeOption, categoriesOption, locationOption, onlineMeetingOption, calendarIdOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var subject = parseResult.GetValue(subjectOption)!;
            var start = parseResult.GetValue(startOption)!;
            var end = parseResult.GetValue(endOption)!;
            var attendees = parseResult.GetValue(attendeesOption);
            var body = parseResult.GetValue(bodyOption);
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";
            var categories = parseResult.GetValue(categoriesOption);
            var location = parseResult.GetValue(locationOption);
            var onlineMeeting = parseResult.GetValue(onlineMeetingOption);
            var calendarId = parseResult.GetValue(calendarIdOption);
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var newEvent = new Event
                {
                    Subject = subject,
                    Start = new DateTimeTimeZone { DateTime = start, TimeZone = tz },
                    End = new DateTimeTimeZone { DateTime = end, TimeZone = tz }
                };

                if (!string.IsNullOrEmpty(body))
                    newEvent.Body = new ItemBody { ContentType = contentType == "html" ? BodyType.Html : BodyType.Text, Content = body };

                if (!string.IsNullOrEmpty(attendees))
                {
                    newEvent.Attendees = attendees.Split(',').Select(e => new Attendee
                    {
                        EmailAddress = new EmailAddress { Address = e.Trim() },
                        Type = AttendeeType.Required
                    }).ToList();
                }

                if (!string.IsNullOrEmpty(categories))
                    newEvent.Categories = categories.Split(',').Select(c => c.Trim()).ToList();

                if (!string.IsNullOrEmpty(location))
                    newEvent.Location = new Location { DisplayName = location };

                if (onlineMeeting)
                {
                    newEvent.IsOnlineMeeting = true;
                    newEvent.OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness;
                }

                Event? created;
                if (!string.IsNullOrEmpty(calendarId))
                    created = await client.Me.Calendars[calendarId].Events.PostAsync(newEvent, cancellationToken: ct);
                else
                    created = await client.Me.Events.PostAsync(newEvent, cancellationToken: ct);

                OutputService.Print(new { status = "created", id = created?.Id, subject, start, end, joinUrl = created?.OnlineMeeting?.JoinUrl });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildUpdateEvent(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var eventIdArg = new Argument<string>("event-id") { Description = "Event ID" };
        var subjectOption = new Option<string?>("--subject") { Description = "New subject" };
        var startOption = new Option<string?>("--start") { Description = "New start datetime" };
        var endOption = new Option<string?>("--end") { Description = "New end datetime" };
        var bodyOption = new Option<string?>("--body") { Description = "New body" };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Body content type: text or html" };
        var categoriesOption = new Option<string?>("--categories") { Description = "Comma-separated category names" };
        var cmd = new Command("update-event", "Update a calendar event") { eventIdArg, subjectOption, startOption, endOption, bodyOption, contentTypeOption, categoriesOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var eventId = parseResult.GetValue(eventIdArg)!;
            var subject = parseResult.GetValue(subjectOption);
            var start = parseResult.GetValue(startOption);
            var end = parseResult.GetValue(endOption);
            var body = parseResult.GetValue(bodyOption);
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";
            var categories = parseResult.GetValue(categoriesOption);
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var update = new Event();

                if (subject != null) update.Subject = subject;
                if (start != null) update.Start = new DateTimeTimeZone { DateTime = start, TimeZone = tz };
                if (end != null) update.End = new DateTimeTimeZone { DateTime = end, TimeZone = tz };
                if (body != null) update.Body = new ItemBody { ContentType = contentType == "html" ? BodyType.Html : BodyType.Text, Content = body };
                if (categories != null) update.Categories = categories.Split(',').Select(c => c.Trim()).ToList();

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

    private static Command BuildFindTimes(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var attendeesOption = new Option<string>("--attendees") { Description = "Comma-separated attendee emails", Required = true };
        var durationOption = new Option<int>("--duration") { Description = "Meeting duration in minutes", Required = true };
        var startOption = new Option<string?>("--start") { Description = "Search window start (ISO 8601, default: now)" };
        var endOption = new Option<string?>("--end") { Description = "Search window end (ISO 8601, default: +7 days)" };
        var cmd = new Command("find-times", "Find available meeting times for attendees") { attendeesOption, durationOption, startOption, endOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var attendees = parseResult.GetValue(attendeesOption)!;
            var duration = parseResult.GetValue(durationOption);
            var start = parseResult.GetValue(startOption) ?? DateTime.Now.ToString("o");
            var end = parseResult.GetValue(endOption) ?? DateTime.Now.AddDays(7).ToString("o");
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var result = await client.Me.FindMeetingTimes.PostAsync(new FindMeetingTimesPostRequestBody
                {
                    Attendees = attendees.Split(',').Select(e => new AttendeeBase
                    {
                        EmailAddress = new EmailAddress { Address = e.Trim() },
                        Type = AttendeeType.Required
                    }).ToList(),
                    TimeConstraint = new TimeConstraint
                    {
                        TimeSlots = [new TimeSlot
                        {
                            Start = new DateTimeTimeZone { DateTime = start, TimeZone = tz },
                            End = new DateTimeTimeZone { DateTime = end, TimeZone = tz }
                        }]
                    },
                    MeetingDuration = XmlConvert.ToTimeSpan($"PT{duration}M"),
                    ReturnSuggestionReasons = true
                }, cancellationToken: ct);

                var suggestions = result?.MeetingTimeSuggestions?.Select(s => new
                {
                    StartDateTime = s.MeetingTimeSlot?.Start?.DateTime,
                    StartTimeZone = s.MeetingTimeSlot?.Start?.TimeZone,
                    EndDateTime = s.MeetingTimeSlot?.End?.DateTime,
                    EndTimeZone = s.MeetingTimeSlot?.End?.TimeZone,
                    Confidence = s.Confidence,
                    OrganizerAvailability = s.OrganizerAvailability?.ToString(),
                    SuggestionReason = s.SuggestionReason
                }).ToList();

                if (suggestions == null || suggestions.Count == 0)
                {
                    OutputService.Print(new { status = "no_suggestions", reason = result?.EmptySuggestionsReason ?? "unknown" }, format);
                }
                else
                {
                    OutputService.Print(suggestions, format);
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

    private static Command BuildSchedule(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var usersOption = new Option<string>("--users") { Description = "Comma-separated user emails", Required = true };
        var startOption = new Option<string>("--start") { Description = "Start datetime (ISO 8601)", Required = true };
        var endOption = new Option<string>("--end") { Description = "End datetime (ISO 8601)", Required = true };
        var cmd = new Command("schedule", "Get free/busy schedule for users") { usersOption, startOption, endOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var users = parseResult.GetValue(usersOption)!;
            var start = parseResult.GetValue(startOption)!;
            var end = parseResult.GetValue(endOption)!;
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var result = await client.Me.Calendar.GetSchedule.PostAsGetSchedulePostResponseAsync(
                    new GetSchedulePostRequestBody
                    {
                        Schedules = users.Split(',').Select(e => e.Trim()).ToList(),
                        StartTime = new DateTimeTimeZone { DateTime = start, TimeZone = tz },
                        EndTime = new DateTimeTimeZone { DateTime = end, TimeZone = tz }
                    }, cancellationToken: ct);

                var schedules = result?.Value?.Select(s => new
                {
                    User = s.ScheduleId,
                    AvailabilityView = s.AvailabilityView,
                    Items = s.ScheduleItems?.Select(i => new
                    {
                        Status = i.Status?.ToString(),
                        Subject = i.Subject,
                        Location = i.Location,
                        StartDateTime = i.Start?.DateTime,
                        StartTimeZone = i.Start?.TimeZone,
                        EndDateTime = i.End?.DateTime,
                        EndTimeZone = i.End?.TimeZone,
                        i.IsPrivate
                    }).ToList()
                }).ToList();

                OutputService.Print(schedules, format);
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
