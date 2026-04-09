using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class CalendarTools
{
    [McpServerTool(Name = "calendar_list"), Description("List all calendars for the current user")]
    public static async Task<string> List()
    {
        try
        {
            var result = await CalendarService.ListAsync();
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "calendar_events"), Description("List calendar events within a date range")]
    public static async Task<string> Events(
        [Description("Start date/time in ISO 8601 format (default: today)")] string? start = null,
        [Description("End date/time in ISO 8601 format (default: +7 days)")] string? end = null,
        [Description("Specific calendar ID (default: primary calendar)")] string? calendarId = null,
        [Description("Maximum number of events to return (default: 25)")] int top = 25,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var result = await CalendarService.EventsAsync(start, end, calendarId, top, timezone);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "calendar_get_event"), Description("Get full details of a specific calendar event including attendees, body, and recurrence")]
    public static async Task<string> GetEvent(
        [Description("The event ID")] string eventId,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var result = await CalendarService.GetEventAsync(eventId, timezone);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "calendar_create_event"), Description("Create a new calendar event")]
    public static async Task<string> CreateEvent(
        [Description("Event subject")] string subject,
        [Description("Start date/time in ISO 8601 format")] string start,
        [Description("End date/time in ISO 8601 format")] string end,
        [Description("Comma-separated attendee email addresses")] string? attendees = null,
        [Description("Event body/description")] string? body = null,
        [Description("Body content type: text or html (default: text)")] string? contentType = null,
        [Description("Comma-separated category names")] string? categories = null,
        [Description("Event location (e.g. room name or address)")] string? location = null,
        [Description("Generate a Teams online meeting link")] bool onlineMeeting = false,
        [Description("Calendar ID to create the event in (default: primary)")] string? calendarId = null,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var result = await CalendarService.CreateEventAsync(
                subject, start, end, attendees, body, contentType,
                categories, location, onlineMeeting, calendarId, timezone);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "calendar_update_event"), Description("Update an existing calendar event")]
    public static async Task<string> UpdateEvent(
        [Description("The event ID")] string eventId,
        [Description("New event subject")] string? subject = null,
        [Description("New start date/time in ISO 8601 format")] string? start = null,
        [Description("New end date/time in ISO 8601 format")] string? end = null,
        [Description("New event body/description")] string? body = null,
        [Description("Body content type: text or html (default: text)")] string? contentType = null,
        [Description("Comma-separated category names")] string? categories = null,
        [Description("Apply update to entire recurring series (resolves series master automatically)")] bool series = false,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var result = await CalendarService.UpdateEventAsync(
                eventId, subject, start, end, body, contentType, categories, series, timezone);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "calendar_delete_event"), Description("Delete a calendar event")]
    public static async Task<string> DeleteEvent(
        [Description("The event ID")] string eventId)
    {
        try
        {
            var result = await CalendarService.DeleteEventAsync(eventId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "calendar_respond"), Description("Respond to a meeting invitation (accept, decline, or tentative)")]
    public static async Task<string> Respond(
        [Description("The event ID")] string eventId,
        [Description("Response action: accept, decline, or tentative")] string action,
        [Description("Optional response comment")] string? comment = null)
    {
        try
        {
            var result = await CalendarService.RespondAsync(eventId, action, comment);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "calendar_find_times"), Description("Find available meeting times for a set of attendees")]
    public static async Task<string> FindTimes(
        [Description("Comma-separated attendee email addresses")] string attendees,
        [Description("Meeting duration in minutes")] int duration,
        [Description("Search window start date/time in ISO 8601 format (default: now)")] string? start = null,
        [Description("Search window end date/time in ISO 8601 format (default: +7 days)")] string? end = null,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var result = await CalendarService.FindTimesAsync(attendees, duration, start, end, timezone);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "calendar_schedule"), Description("Get free/busy schedule for one or more users")]
    public static async Task<string> Schedule(
        [Description("Comma-separated user email addresses")] string users,
        [Description("Start date/time in ISO 8601 format")] string start,
        [Description("End date/time in ISO 8601 format")] string end,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var result = await CalendarService.ScheduleAsync(users, start, end, timezone);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
