using System.CommandLine;
using GraphCli.Services;
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
                var result = await CalendarService.ListAsync();
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
            var start = parseResult.GetValue(startOption);
            var end = parseResult.GetValue(endOption);
            var calendarId = parseResult.GetValue(calendarIdOption);
            var top = parseResult.GetValue(topOption);
            var tz = parseResult.GetValue(timezoneOption);

            try
            {
                var result = await CalendarService.EventsAsync(start, end, calendarId, top, tz);
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

    private static Command BuildGetEvent(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var eventIdArg = new Argument<string>("event-id") { Description = "Event ID" };
        var cmd = new Command("get-event", "Get full event details including attendees and body") { eventIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var eventId = parseResult.GetValue(eventIdArg)!;
            var tz = parseResult.GetValue(timezoneOption);
            try
            {
                var result = await CalendarService.GetEventAsync(eventId, tz);
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
            var tz = parseResult.GetValue(timezoneOption);

            try
            {
                var result = await CalendarService.CreateEventAsync(
                    subject, start, end, attendees, body, contentType,
                    categories, location, onlineMeeting, calendarId, tz);
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

    private static Command BuildUpdateEvent(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var eventIdArg = new Argument<string>("event-id") { Description = "Event ID" };
        var subjectOption = new Option<string?>("--subject") { Description = "New subject" };
        var startOption = new Option<string?>("--start") { Description = "New start datetime" };
        var endOption = new Option<string?>("--end") { Description = "New end datetime" };
        var bodyOption = new Option<string?>("--body") { Description = "New body" };
        var contentTypeOption = new Option<string>("--content-type") { DefaultValueFactory = _ => "text", Description = "Body content type: text or html" };
        var categoriesOption = new Option<string?>("--categories") { Description = "Comma-separated category names" };
        var seriesOption = new Option<bool>("--series") { Description = "Apply update to the entire series (resolves series master automatically)" };
        var cmd = new Command("update-event", "Update a calendar event") { eventIdArg, subjectOption, startOption, endOption, bodyOption, contentTypeOption, categoriesOption, seriesOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var eventId = parseResult.GetValue(eventIdArg)!;
            var subject = parseResult.GetValue(subjectOption);
            var start = parseResult.GetValue(startOption);
            var end = parseResult.GetValue(endOption);
            var body = parseResult.GetValue(bodyOption);
            var contentType = parseResult.GetValue(contentTypeOption) ?? "text";
            var categories = parseResult.GetValue(categoriesOption);
            var series = parseResult.GetValue(seriesOption);
            var tz = parseResult.GetValue(timezoneOption);

            try
            {
                var result = await CalendarService.UpdateEventAsync(
                    eventId, subject, start, end, body, contentType, categories, series, tz);
                OutputService.Print(result);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
            catch (InvalidOperationException ex)
            {
                OutputService.PrintError("not_recurring", ex.Message);
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
                var result = await CalendarService.DeleteEventAsync(eventId);
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

    private static Command BuildRespond(Option<string> formatOption)
    {
        var eventIdArg = new Argument<string>("event-id") { Description = "Event ID" };
        var actionOption = new Option<string>("--action") { Description = "Response: accept, decline, or tentative", Required = true };
        var commentOption = new Option<string?>("--comment") { Description = "Optional response comment" };
        var cmd = new Command("respond", "Respond to a meeting invitation") { eventIdArg, actionOption, commentOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var eventId = parseResult.GetValue(eventIdArg)!;
            var action = parseResult.GetValue(actionOption)!;
            var comment = parseResult.GetValue(commentOption);

            try
            {
                var result = await CalendarService.RespondAsync(eventId, action, comment);
                OutputService.Print(result);
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
            catch (InvalidOperationException ex)
            {
                OutputService.PrintError("invalid_action", ex.Message);
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
            var start = parseResult.GetValue(startOption);
            var end = parseResult.GetValue(endOption);
            var tz = parseResult.GetValue(timezoneOption);

            try
            {
                var result = await CalendarService.FindTimesAsync(attendees, duration, start, end, tz);
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
            var tz = parseResult.GetValue(timezoneOption);

            try
            {
                var result = await CalendarService.ScheduleAsync(users, start, end, tz);
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
}
