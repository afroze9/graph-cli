using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class TaskCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var taskCommand = new Command("task", "Microsoft To Do task operations");

        taskCommand.Subcommands.Add(BuildLists(formatOption));
        taskCommand.Subcommands.Add(BuildList(formatOption));
        taskCommand.Subcommands.Add(BuildCreate(formatOption));
        taskCommand.Subcommands.Add(BuildUpdate(formatOption));
        taskCommand.Subcommands.Add(BuildDelete(formatOption));
        taskCommand.Subcommands.Add(BuildComplete(formatOption));

        return taskCommand;
    }

    private static Command BuildLists(Option<string> formatOption)
    {
        var cmd = new Command("lists", "List all task lists");
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var lists = await client.Me.Todo.Lists.GetAsync(cancellationToken: ct);
                var results = lists?.Value?.Select(l => new
                {
                    l.Id,
                    l.DisplayName,
                    l.IsOwner,
                    l.IsShared,
                    WellknownListName = l.WellknownListName?.ToString()
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

    private static Command BuildList(Option<string> formatOption)
    {
        var listIdArg = new Argument<string>("list-id") { Description = "Task list ID" };
        var statusOption = new Option<string?>("--status") { Description = "Filter by status: notStarted, inProgress, completed" };
        var cmd = new Command("list", "Get tasks from a list") { listIdArg, statusOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var listId = parseResult.GetValue(listIdArg)!;
            var status = parseResult.GetValue(statusOption);
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var tasks = await client.Me.Todo.Lists[listId].Tasks.GetAsync(r =>
                {
                    if (!string.IsNullOrEmpty(status))
                        r.QueryParameters.Filter = $"status eq '{status}'";
                }, ct);
                var results = tasks?.Value?.Select(t => new
                {
                    t.Id,
                    t.Title,
                    Status = t.Status?.ToString(),
                    Importance = t.Importance?.ToString(),
                    DueDate = t.DueDateTime?.DateTime,
                    DueTimeZone = t.DueDateTime?.TimeZone,
                    t.CreatedDateTime,
                    t.LastModifiedDateTime,
                    t.CompletedDateTime
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

    private static Command BuildCreate(Option<string> formatOption)
    {
        var listIdArg = new Argument<string>("list-id") { Description = "Task list ID" };
        var titleOption = new Option<string>("--title") { Description = "Task title", Required = true };
        var dueOption = new Option<string?>("--due") { Description = "Due date (ISO 8601)" };
        var importanceOption = new Option<string?>("--importance") { Description = "Importance: low, normal, or high" };
        var bodyOption = new Option<string?>("--body") { Description = "Task body/notes" };
        var cmd = new Command("create", "Create a task") { listIdArg, titleOption, dueOption, importanceOption, bodyOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var listId = parseResult.GetValue(listIdArg)!;
            var title = parseResult.GetValue(titleOption)!;
            var due = parseResult.GetValue(dueOption);
            var importance = parseResult.GetValue(importanceOption);
            var body = parseResult.GetValue(bodyOption);

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var task = new TodoTask
                {
                    Title = title
                };

                if (!string.IsNullOrEmpty(due))
                    task.DueDateTime = new DateTimeTimeZone { DateTime = due, TimeZone = TimeZoneInfo.Local.Id };

                if (!string.IsNullOrEmpty(importance))
                {
                    task.Importance = importance.ToLower() switch
                    {
                        "low" => Microsoft.Graph.Models.Importance.Low,
                        "high" => Microsoft.Graph.Models.Importance.High,
                        _ => Microsoft.Graph.Models.Importance.Normal
                    };
                }

                if (!string.IsNullOrEmpty(body))
                    task.Body = new ItemBody { ContentType = BodyType.Text, Content = body };

                var created = await client.Me.Todo.Lists[listId].Tasks.PostAsync(task, cancellationToken: ct);
                OutputService.Print(new { status = "created", id = created?.Id, title });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildUpdate(Option<string> formatOption)
    {
        var listIdArg = new Argument<string>("list-id") { Description = "Task list ID" };
        var taskIdArg = new Argument<string>("task-id") { Description = "Task ID" };
        var titleOption = new Option<string?>("--title") { Description = "New title" };
        var statusOption = new Option<string?>("--status") { Description = "New status: notStarted, inProgress, completed" };
        var dueOption = new Option<string?>("--due") { Description = "New due date" };
        var importanceOption = new Option<string?>("--importance") { Description = "New importance: low, normal, high" };
        var cmd = new Command("update", "Update a task") { listIdArg, taskIdArg, titleOption, statusOption, dueOption, importanceOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var listId = parseResult.GetValue(listIdArg)!;
            var taskId = parseResult.GetValue(taskIdArg)!;
            var title = parseResult.GetValue(titleOption);
            var statusStr = parseResult.GetValue(statusOption);
            var due = parseResult.GetValue(dueOption);
            var importance = parseResult.GetValue(importanceOption);

            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var update = new TodoTask();

                if (title != null) update.Title = title;
                if (due != null) update.DueDateTime = new DateTimeTimeZone { DateTime = due, TimeZone = TimeZoneInfo.Local.Id };

                if (statusStr != null)
                {
                    update.Status = statusStr.ToLower() switch
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
                        "low" => Microsoft.Graph.Models.Importance.Low,
                        "high" => Microsoft.Graph.Models.Importance.High,
                        _ => Microsoft.Graph.Models.Importance.Normal
                    };
                }

                var updated = await client.Me.Todo.Lists[listId].Tasks[taskId].PatchAsync(update, cancellationToken: ct);
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

    private static Command BuildDelete(Option<string> formatOption)
    {
        var listIdArg = new Argument<string>("list-id") { Description = "Task list ID" };
        var taskIdArg = new Argument<string>("task-id") { Description = "Task ID" };
        var cmd = new Command("delete", "Delete a task") { listIdArg, taskIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var listId = parseResult.GetValue(listIdArg)!;
            var taskId = parseResult.GetValue(taskIdArg)!;
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                await client.Me.Todo.Lists[listId].Tasks[taskId].DeleteAsync(cancellationToken: ct);
                OutputService.Print(new { status = "deleted", listId, taskId });
            }
            catch (ODataError ex)
            {
                OutputService.PrintError(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return cmd;
    }

    private static Command BuildComplete(Option<string> formatOption)
    {
        var listIdArg = new Argument<string>("list-id") { Description = "Task list ID" };
        var taskIdArg = new Argument<string>("task-id") { Description = "Task ID" };
        var cmd = new Command("complete", "Mark a task as completed") { listIdArg, taskIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var listId = parseResult.GetValue(listIdArg)!;
            var taskId = parseResult.GetValue(taskIdArg)!;
            try
            {
                var client = await GraphClientProvider.CreateAsync();
                var update = new TodoTask { Status = Microsoft.Graph.Models.TaskStatus.Completed };
                var updated = await client.Me.Todo.Lists[listId].Tasks[taskId].PatchAsync(update, cancellationToken: ct);
                OutputService.Print(new { status = "completed", id = updated?.Id });
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
