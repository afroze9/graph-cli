using Microsoft.Graph.Models;

namespace GraphCli.Services;

public static class TaskService
{
    public static async Task<object> ListTaskListsAsync()
    {
        var client = await GraphClientProvider.CreateAsync();
        var lists = await client.Me.Todo.Lists.GetAsync();
        return lists?.Value?.Select(l => new
        {
            l.Id,
            l.DisplayName,
            l.IsOwner,
            l.IsShared,
            WellknownListName = l.WellknownListName?.ToString()
        }).ToList() ?? [];
    }

    public static async Task<object> ListTasksAsync(string listId, string? status, string tz)
    {
        var client = await GraphClientProvider.CreateAsync();
        var tasks = await client.Me.Todo.Lists[listId].Tasks.GetAsync(r =>
        {
            if (!string.IsNullOrEmpty(status))
                r.QueryParameters.Filter = $"status eq '{status}'";
        });
        return tasks?.Value?.Select(t => new
        {
            t.Id,
            t.Title,
            Status = t.Status?.ToString(),
            Importance = t.Importance?.ToString(),
            DueDate = TimeZoneService.ConvertToTimeZone(t.DueDateTime?.DateTime, t.DueDateTime?.TimeZone, tz),
            DueTimeZone = tz,
            CreatedDateTime = TimeZoneService.ConvertToTimeZone(t.CreatedDateTime, tz),
            LastModifiedDateTime = TimeZoneService.ConvertToTimeZone(t.LastModifiedDateTime, tz),
            CompletedDateTime = TimeZoneService.ConvertToTimeZone(t.CompletedDateTime?.DateTime, t.CompletedDateTime?.TimeZone, tz)
        }).ToList() ?? [];
    }

    public static async Task<object> CreateTaskAsync(string listId, string title, string? due, string? importance, string? body, string tz)
    {
        var client = await GraphClientProvider.CreateAsync();
        var task = new TodoTask
        {
            Title = title
        };

        if (!string.IsNullOrEmpty(due))
            task.DueDateTime = new DateTimeTimeZone { DateTime = due, TimeZone = tz };

        if (!string.IsNullOrEmpty(importance))
        {
            task.Importance = importance.ToLower() switch
            {
                "low" => Importance.Low,
                "high" => Importance.High,
                _ => Importance.Normal
            };
        }

        if (!string.IsNullOrEmpty(body))
            task.Body = new ItemBody { ContentType = BodyType.Text, Content = body };

        var created = await client.Me.Todo.Lists[listId].Tasks.PostAsync(task);
        return new { status = "created", id = created?.Id, title };
    }

    public static async Task<object> UpdateTaskAsync(string listId, string taskId, string? title, string? status, string? due, string? importance, string tz)
    {
        var client = await GraphClientProvider.CreateAsync();
        var update = new TodoTask();

        if (title != null) update.Title = title;
        if (due != null) update.DueDateTime = new DateTimeTimeZone { DateTime = due, TimeZone = tz };

        if (status != null)
        {
            update.Status = status.ToLower() switch
            {
                "notstarted" => Microsoft.Graph.Models.TaskStatus.NotStarted,
                "inprogress" => Microsoft.Graph.Models.TaskStatus.InProgress,
                "completed" => Microsoft.Graph.Models.TaskStatus.Completed,
                _ => Microsoft.Graph.Models.TaskStatus.NotStarted
            };
        }

        if (importance != null)
        {
            update.Importance = importance.ToLower() switch
            {
                "low" => Importance.Low,
                "high" => Importance.High,
                _ => Importance.Normal
            };
        }

        var updated = await client.Me.Todo.Lists[listId].Tasks[taskId].PatchAsync(update);
        return new { status = "updated", id = updated?.Id };
    }

    public static async Task<object> DeleteTaskAsync(string listId, string taskId)
    {
        var client = await GraphClientProvider.CreateAsync();
        await client.Me.Todo.Lists[listId].Tasks[taskId].DeleteAsync();
        return new { status = "deleted", listId, taskId };
    }

    public static async Task<object> CompleteTaskAsync(string listId, string taskId)
    {
        var client = await GraphClientProvider.CreateAsync();
        var update = new TodoTask { Status = Microsoft.Graph.Models.TaskStatus.Completed };
        var updated = await client.Me.Todo.Lists[listId].Tasks[taskId].PatchAsync(update);
        return new { status = "completed", id = updated?.Id };
    }
}
