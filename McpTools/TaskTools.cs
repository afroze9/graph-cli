using System.ComponentModel;
using GraphCli.Services;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class TaskTools
{
    [McpServerTool(Name = "task_lists"), Description("List all Microsoft To Do task lists")]
    public static async Task<string> Lists()
    {
        try
        {
            var result = await TaskService.ListTaskListsAsync();
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "task_list"), Description("Get tasks from a specific task list")]
    public static async Task<string> List(
        [Description("Task list ID")] string listId,
        [Description("Filter by status: notStarted, inProgress, completed")] string? status = null,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var tz = TimeZoneService.ResolveTimeZoneId(timezone);
            var result = await TaskService.ListTasksAsync(listId, status, tz);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "task_create"), Description("Create a new task in a task list")]
    public static async Task<string> Create(
        [Description("Task list ID")] string listId,
        [Description("Task title")] string title,
        [Description("Due date (ISO 8601)")] string? due = null,
        [Description("Importance: low, normal, or high")] string? importance = null,
        [Description("Task body/notes")] string? body = null,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var tz = TimeZoneService.ResolveTimeZoneId(timezone);
            var result = await TaskService.CreateTaskAsync(listId, title, due, importance, body, tz);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "task_update"), Description("Update an existing task")]
    public static async Task<string> Update(
        [Description("Task list ID")] string listId,
        [Description("Task ID")] string taskId,
        [Description("New title")] string? title = null,
        [Description("New status: notStarted, inProgress, completed")] string? status = null,
        [Description("New due date (ISO 8601)")] string? due = null,
        [Description("New importance: low, normal, high")] string? importance = null,
        [Description("IANA timezone (e.g. Asia/Karachi)")] string? timezone = null)
    {
        try
        {
            var tz = TimeZoneService.ResolveTimeZoneId(timezone);
            var result = await TaskService.UpdateTaskAsync(listId, taskId, title, status, due, importance, tz);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "task_delete"), Description("Delete a task")]
    public static async Task<string> Delete(
        [Description("Task list ID")] string listId,
        [Description("Task ID")] string taskId)
    {
        try
        {
            var result = await TaskService.DeleteTaskAsync(listId, taskId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }

    [McpServerTool(Name = "task_complete"), Description("Mark a task as completed")]
    public static async Task<string> Complete(
        [Description("Task list ID")] string listId,
        [Description("Task ID")] string taskId)
    {
        try
        {
            var result = await TaskService.CompleteTaskAsync(listId, taskId);
            return McpGraphHelper.ToJson(result);
        }
        catch (ODataError ex) { return McpGraphHelper.HandleODataError(ex); }
        catch (Exception ex) { return McpGraphHelper.HandleException(ex); }
    }
}
