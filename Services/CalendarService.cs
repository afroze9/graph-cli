using System.Xml;
using Microsoft.Graph.Me.Calendar.GetSchedule;
using Microsoft.Graph.Me.FindMeetingTimes;
using Microsoft.Graph.Models;

namespace GraphCli.Services;

public static class CalendarService
{
    public static async Task<object> ListAsync()
    {
        var client = await GraphClientProvider.CreateAsync();
        var calendars = await client.Me.Calendars.GetAsync(r =>
        {
            r.QueryParameters.Select = ["id", "name", "color", "isDefaultCalendar", "canEdit", "owner"];
        });
        return calendars?.Value?.Select(c => new
        {
            c.Id,
            c.Name,
            Color = c.Color?.ToString(),
            c.IsDefaultCalendar,
            c.CanEdit,
            OwnerName = c.Owner?.Name,
            OwnerEmail = c.Owner?.Address
        }).ToList() ?? [];
    }

    public static async Task<object> EventsAsync(
        string? start = null,
        string? end = null,
        string? calendarId = null,
        int top = 25,
        string? tz = null)
    {
        var client = await GraphClientProvider.CreateAsync();
        tz = TimeZoneService.ResolveTimeZoneId(tz);
        var startDt = start ?? DateTime.Today.ToString("o");
        var endDt = end ?? DateTime.Today.AddDays(7).ToString("o");
        string[] select = ["id", "subject", "start", "end", "location", "organizer", "isAllDay", "isCancelled", "responseStatus", "categories", "type", "seriesMasterId"];

        EventCollectionResponse? events;
        if (!string.IsNullOrEmpty(calendarId))
        {
            events = await client.Me.Calendars[calendarId].CalendarView.GetAsync(r =>
            {
                r.QueryParameters.StartDateTime = startDt;
                r.QueryParameters.EndDateTime = endDt;
                r.QueryParameters.Top = top;
                r.QueryParameters.Select = select;
                r.QueryParameters.Orderby = ["start/dateTime"];
                r.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
            });
        }
        else
        {
            events = await client.Me.CalendarView.GetAsync(r =>
            {
                r.QueryParameters.StartDateTime = startDt;
                r.QueryParameters.EndDateTime = endDt;
                r.QueryParameters.Top = top;
                r.QueryParameters.Select = select;
                r.QueryParameters.Orderby = ["start/dateTime"];
                r.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
            });
        }

        return events?.Value?.Select(e => new
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
            Categories = e.Categories,
            Type = e.Type?.ToString(),
            e.SeriesMasterId
        }).ToList() ?? [];
    }

    public static async Task<object> GetEventAsync(string eventId, string? tz = null)
    {
        var client = await GraphClientProvider.CreateAsync();
        tz = TimeZoneService.ResolveTimeZoneId(tz);
        var e = await client.Me.Events[eventId].GetAsync(r =>
        {
            r.QueryParameters.Select =
            [
                "id", "subject", "body", "bodyPreview", "start", "end",
                "location", "locations", "organizer", "attendees",
                "isOnlineMeeting", "onlineMeeting", "onlineMeetingProvider",
                "importance", "sensitivity", "isAllDay", "isCancelled",
                "responseStatus", "categories", "hasAttachments",
                "recurrence", "type", "seriesMasterId", "webLink"
            ];
            r.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
        });

        return new
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
            Type = e.Type?.ToString(),
            e.SeriesMasterId,
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
        };
    }

    public static async Task<object> CreateEventAsync(
        string subject,
        string start,
        string end,
        string? attendees = null,
        string? body = null,
        string? contentType = null,
        string? categories = null,
        string? location = null,
        bool onlineMeeting = false,
        string? calendarId = null,
        string? tz = null)
    {
        var client = await GraphClientProvider.CreateAsync();
        tz = TimeZoneService.ResolveTimeZoneId(tz);

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
            created = await client.Me.Calendars[calendarId].Events.PostAsync(newEvent);
        else
            created = await client.Me.Events.PostAsync(newEvent);

        return new { status = "created", id = created?.Id, subject, start, end, joinUrl = created?.OnlineMeeting?.JoinUrl };
    }

    public static async Task<object> UpdateEventAsync(
        string eventId,
        string? subject = null,
        string? start = null,
        string? end = null,
        string? body = null,
        string? contentType = null,
        string? categories = null,
        bool series = false,
        string? tz = null)
    {
        var client = await GraphClientProvider.CreateAsync();
        tz = TimeZoneService.ResolveTimeZoneId(tz);

        if (series)
        {
            var existing = await client.Me.Events[eventId].GetAsync(r =>
            {
                r.QueryParameters.Select = ["type", "seriesMasterId"];
            });
            var eventType = existing?.Type?.ToString();
            if (eventType == "Occurrence" || eventType == "Exception")
                eventId = existing!.SeriesMasterId!;
            else if (eventType != "SeriesMaster")
                throw new InvalidOperationException("Event is not part of a recurring series");
        }

        var update = new Event();

        if (subject != null) update.Subject = subject;
        if (start != null) update.Start = new DateTimeTimeZone { DateTime = start, TimeZone = tz };
        if (end != null) update.End = new DateTimeTimeZone { DateTime = end, TimeZone = tz };
        if (body != null) update.Body = new ItemBody { ContentType = contentType == "html" ? BodyType.Html : BodyType.Text, Content = body };
        if (categories != null) update.Categories = categories.Split(',').Select(c => c.Trim()).ToList();

        var updated = await client.Me.Events[eventId].PatchAsync(update);
        return new { status = "updated", id = updated?.Id, series };
    }

    public static async Task<object> DeleteEventAsync(string eventId)
    {
        var client = await GraphClientProvider.CreateAsync();
        await client.Me.Events[eventId].DeleteAsync();
        return new { status = "deleted", eventId };
    }

    public static async Task<object> RespondAsync(string eventId, string action, string? comment = null)
    {
        var client = await GraphClientProvider.CreateAsync();

        switch (action.ToLowerInvariant())
        {
            case "accept":
                await client.Me.Events[eventId].Accept.PostAsync(
                    new Microsoft.Graph.Me.Events.Item.Accept.AcceptPostRequestBody
                    {
                        Comment = comment,
                        SendResponse = true
                    });
                break;
            case "decline":
                await client.Me.Events[eventId].Decline.PostAsync(
                    new Microsoft.Graph.Me.Events.Item.Decline.DeclinePostRequestBody
                    {
                        Comment = comment,
                        SendResponse = true
                    });
                break;
            case "tentative":
                await client.Me.Events[eventId].TentativelyAccept.PostAsync(
                    new Microsoft.Graph.Me.Events.Item.TentativelyAccept.TentativelyAcceptPostRequestBody
                    {
                        Comment = comment,
                        SendResponse = true
                    });
                break;
            default:
                throw new InvalidOperationException($"Unknown action: {action}. Use 'accept', 'decline', or 'tentative'.");
        }

        return new { status = "responded", eventId, action };
    }

    public static async Task<object> FindTimesAsync(
        string attendees,
        int duration,
        string? start = null,
        string? end = null,
        string? tz = null)
    {
        var client = await GraphClientProvider.CreateAsync();
        tz = TimeZoneService.ResolveTimeZoneId(tz);
        var startDt = start ?? DateTime.Now.ToString("o");
        var endDt = end ?? DateTime.Now.AddDays(7).ToString("o");

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
                    Start = new DateTimeTimeZone { DateTime = startDt, TimeZone = tz },
                    End = new DateTimeTimeZone { DateTime = endDt, TimeZone = tz }
                }]
            },
            MeetingDuration = XmlConvert.ToTimeSpan($"PT{duration}M"),
            ReturnSuggestionReasons = true
        });

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
            return new { status = "no_suggestions", reason = result?.EmptySuggestionsReason ?? "unknown" };

        return suggestions;
    }

    public static async Task<object> ScheduleAsync(
        string users,
        string start,
        string end,
        string? tz = null)
    {
        var client = await GraphClientProvider.CreateAsync();
        tz = TimeZoneService.ResolveTimeZoneId(tz);

        var result = await client.Me.Calendar.GetSchedule.PostAsGetSchedulePostResponseAsync(
            new GetSchedulePostRequestBody
            {
                Schedules = users.Split(',').Select(e => e.Trim()).ToList(),
                StartTime = new DateTimeTimeZone { DateTime = start, TimeZone = tz },
                EndTime = new DateTimeTimeZone { DateTime = end, TimeZone = tz }
            });

        return result?.Value?.Select(s => new
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
        }).ToList() ?? [];
    }
}
