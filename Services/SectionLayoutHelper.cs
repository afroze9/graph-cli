namespace GraphCli.Services;

public static class SectionLayoutHelper
{
    public record SectionLayoutInfo(string Name, string Description, (string Id, int Width)[] Columns);

    private static readonly SectionLayoutInfo[] Layouts =
    [
        new("fullWidth", "Single full-width column (no side margins)", [("1", 0)]),
        new("oneColumn", "Single column (12-grid width)", [("1", 12)]),
        new("twoColumns", "Two equal columns (6+6)", [("1", 6), ("2", 6)]),
        new("threeColumns", "Three equal columns (4+4+4)", [("1", 4), ("2", 4), ("3", 4)]),
        new("oneThirdLeftColumn", "Narrow left + wide right (4+8)", [("1", 4), ("2", 8)]),
        new("oneThirdRightColumn", "Wide left + narrow right (8+4)", [("1", 8), ("2", 4)])
    ];

    private static readonly string[] EmphasisValues = ["none", "netural", "soft", "strong"];

    public static IReadOnlyList<SectionLayoutInfo> All() => Layouts;

    public static (string Id, int Width)[]? GetColumns(string layout)
    {
        var info = Layouts.FirstOrDefault(l =>
            l.Name.Equals(layout, StringComparison.OrdinalIgnoreCase));
        return info?.Columns;
    }

    public static string? ValidateLayout(string layout)
    {
        if (Layouts.Any(l => l.Name.Equals(layout, StringComparison.OrdinalIgnoreCase)))
            return null;
        return $"Invalid layout '{layout}'. Valid values: {string.Join(", ", Layouts.Select(l => l.Name))}";
    }

    public static string? ValidateEmphasis(string emphasis)
    {
        if (EmphasisValues.Contains(emphasis, StringComparer.OrdinalIgnoreCase))
            return null;
        return $"Invalid emphasis '{emphasis}'. Valid values: {string.Join(", ", EmphasisValues)}";
    }
}
