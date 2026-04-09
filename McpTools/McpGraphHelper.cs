using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.McpTools;

public static class McpGraphHelper
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static string ToJson(object? data) =>
        JsonSerializer.Serialize(data, JsonOptions);

    public static string Error(string code, string message)
    {
        Console.Error.WriteLine($"[graph-cli] {code}: {message}");
        return ToJson(new { error = code, message });
    }

    public static string HandleODataError(ODataError ex)
    {
        Console.Error.WriteLine($"[graph-cli] ODataError: {ex.Error?.Code} - {ex.Error?.Message}");
        Console.Error.WriteLine($"[graph-cli] {ex}");
        return Error(ex.Error?.Code ?? "error", ex.Error?.Message ?? ex.Message);
    }

    public static string HandleException(Exception ex)
    {
        Console.Error.WriteLine($"[graph-cli] {ex.GetType().Name}: {ex.Message}");
        Console.Error.WriteLine($"[graph-cli] {ex}");
        return Error(ex.GetType().Name, ex.Message);
    }
}
