using System.CommandLine;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Commands;

public static class TaskCommands
{
    public static Command Build(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var taskCommand = new Command("task", "Microsoft To Do task operations");

        taskCommand.Subcommands.Add(BuildLists(formatOption));
        taskCommand.Subcommands.Add(BuildList(formatOption, timezoneOption));
        taskCommand.Subcommands.Add(BuildCreate(formatOption, timezoneOption));
        taskCommand.Subcommands.Add(BuildUpdate(formatOption, timezoneOption));
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
                var result = await TaskService.ListTaskListsAsync();
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

    private static Command BuildList(Option<string> formatOption, Option<string?> timezoneOption)
    {
        var listIdArg = new Argument<string>("list-id") { Description = "Task list ID" };
        var statusOption = new Option<string?>("--status") { Description = "Filter by status: notStarted, inProgress, completed" };
        var cmd = new Command("list", "Get tasks from a list") { listIdArg, statusOption };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var listId = parseResult.GetValue(listIdArg)!;
            var status = parseResult.GetValue(statusOption);
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));
            try
            {
                var result = await TaskService.ListTasksAsync(listId, status, tz);
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

    private static Command BuildCreate(Option<string> formatOption, Option<string?> timezoneOption)
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
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));
            try
            {
                var result = await TaskService.CreateTaskAsync(listId, title, due, importance, body, tz);
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

    private static Command BuildUpdate(Option<string> formatOption, Option<string?> timezoneOption)
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
            var tz = TimeZoneService.ResolveTimeZoneId(parseResult.GetValue(timezoneOption));
            try
            {
                var result = await TaskService.UpdateTaskAsync(listId, taskId, title, statusStr, due, importance, tz);
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
        var listIdArg = new Argument<string>("list-id") { Description = "Task list ID" };
        var taskIdArg = new Argument<string>("task-id") { Description = "Task ID" };
        var cmd = new Command("delete", "Delete a task") { listIdArg, taskIdArg };
        cmd.SetAction(async (parseResult, ct) =>
        {
            var listId = parseResult.GetValue(listIdArg)!;
            var taskId = parseResult.GetValue(taskIdArg)!;
            try
            {
                var result = await TaskService.DeleteTaskAsync(listId, taskId);
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
                var result = await TaskService.CompleteTaskAsync(listId, taskId);
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
}
