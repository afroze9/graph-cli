using System.Text.Json;
using System.Text.Json.Serialization;

namespace GraphCli.Services;

public static class OutputService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static void Print(object? data, string format = "json")
    {
        if (data == null)
        {
            Console.WriteLine("null");
            return;
        }

        if (format == "table")
            PrintTable(data);
        else
            PrintJson(data);
    }

    public static void PrintError(string code, string message)
    {
        var error = new { error = code, message };
        Console.Error.WriteLine(JsonSerializer.Serialize(error, JsonOptions));
    }

    private static void PrintJson(object data)
    {
        Console.WriteLine(JsonSerializer.Serialize(data, JsonOptions));
    }

    private static void PrintTable(object data)
    {
        if (data is System.Collections.IEnumerable enumerable and not string)
        {
            var items = enumerable.Cast<object>().ToList();
            if (items.Count == 0)
            {
                Console.WriteLine("(no results)");
                return;
            }

            var props = items[0].GetType().GetProperties();
            var widths = new int[props.Length];
            var headers = props.Select(p => p.Name).ToArray();

            // Calculate column widths
            for (int i = 0; i < props.Length; i++)
                widths[i] = headers[i].Length;

            var rows = new List<string[]>();
            foreach (var item in items)
            {
                var row = new string[props.Length];
                for (int i = 0; i < props.Length; i++)
                {
                    row[i] = props[i].GetValue(item)?.ToString() ?? "";
                    widths[i] = Math.Max(widths[i], row[i].Length);
                }
                rows.Add(row);
            }

            // Print header
            Console.WriteLine(string.Join("  ", headers.Select((h, i) => h.PadRight(widths[i]))));
            Console.WriteLine(string.Join("  ", widths.Select(w => new string('-', w))));

            // Print rows
            foreach (var row in rows)
                Console.WriteLine(string.Join("  ", row.Select((v, i) => v.PadRight(widths[i]))));
        }
        else
        {
            // Single object: key-value pairs
            var props = data.GetType().GetProperties();
            var maxKeyLen = props.Max(p => p.Name.Length);
            foreach (var prop in props)
            {
                var value = prop.GetValue(data)?.ToString() ?? "";
                Console.WriteLine($"{prop.Name.PadRight(maxKeyLen)}  {value}");
            }
        }
    }
}
