using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Serialization;

namespace GraphCli.Services;

public static class WebPartRegistry
{
    public record WebPartTypeInfo(
        string Guid, string Name, string Description,
        string[] RequiredParams, string[] OptionalParams,
        bool IsPassthrough, Func<Dictionary<string, string>, WebPartData>? Builder);

    private static readonly Dictionary<string, WebPartTypeInfo> ByName;
    private static readonly Dictionary<string, WebPartTypeInfo> ByGuid;

    static WebPartRegistry()
    {
        var types = new WebPartTypeInfo[]
        {
            new("0f087d7f-520e-42b7-89c0-496aaf979d58", "button", "A clickable button with a link",
                ["text", "link"], ["alignment"], false, BuildButton),

            new("df8e44e7-edd5-46d5-90da-aca1539313b8", "calltoaction", "Call to action with image, overlay text, and button",
                [], [], true, null),

            new("544dd15b-cf3c-441b-96da-004d5a8cea1d", "youtube", "Embed a YouTube video",
                [], [], true, null),

            new("2161a1c6-db61-4731-b97c-3cdb303f7cbb", "divider", "A horizontal line divider",
                [], [], false, BuildDivider),

            new("8654b779-4886-46d4-8ffb-b5ed960ee986", "spacer", "Vertical space between webparts",
                [], ["height"], false, BuildSpacer),

            new("d1d91016-032f-456d-98a4-721247c305e8", "image", "Display an image",
                ["imageUrl"], ["altText", "caption"], false, BuildImage),

            new("6410b3b6-d440-4663-8744-378976dc041e", "linkpreview", "Preview card for a URL",
                ["link"], [], false, BuildLinkPreview),

            new("b7dd04e1-19ce-4b24-9132-b60a1c2b910d", "documentembed", "Embed a document",
                [], [], true, null),

            new("e377ea37-9047-43b9-8cdb-a761be2f8e09", "bingmaps", "Bing Maps embed",
                [], [], true, null),

            new("af8be689-990e-492a-81f7-ba3e4cd3ed9c", "imagegallery", "Image gallery carousel",
                [], [], true, null),

            new("e84a8ca2-f63c-4fb9-bc0b-d8eef5ccb22b", "orgchart", "Organization chart",
                [], [], true, null),

            new("7f718435-ee4d-431c-bdbf-9c4ff326f46e", "people", "People cards",
                [], [], true, null),

            new("c70391ea-0b10-4ee9-b2b4-006d3fcad0cd", "quicklinks", "Quick links collection",
                [], [], true, null),

            new("cbe7b0a9-3504-44dd-a3a3-0e5cacd07788", "titlearea", "Title area banner",
                [], [], true, null),
        };

        ByName = types.ToDictionary(t => t.Name, StringComparer.OrdinalIgnoreCase);
        ByGuid = types.ToDictionary(t => t.Guid, StringComparer.OrdinalIgnoreCase);
    }

    public static WebPartTypeInfo? Resolve(string nameOrGuid)
    {
        if (ByName.TryGetValue(nameOrGuid, out var byName)) return byName;
        if (ByGuid.TryGetValue(nameOrGuid, out var byGuid)) return byGuid;
        return null;
    }

    public static IReadOnlyList<WebPartTypeInfo> All() =>
        ByName.Values.ToList();

    public static WebPartData BuildData(string nameOrGuid, Dictionary<string, string> parameters)
    {
        var info = Resolve(nameOrGuid)
            ?? throw new InvalidOperationException($"Unknown webpart type '{nameOrGuid}'.");

        if (info.IsPassthrough)
            throw new InvalidOperationException(
                $"Webpart type '{info.Name}' requires --data-json (schema not yet mapped).");

        if (info.Builder == null)
            throw new InvalidOperationException(
                $"No builder registered for webpart type '{info.Name}'.");

        var missing = info.RequiredParams.Where(p => !parameters.ContainsKey(p)).ToArray();
        if (missing.Length > 0)
            throw new InvalidOperationException(
                $"Missing required parameter(s) for {info.Name}: {string.Join(", ", missing)}. " +
                $"Required: {string.Join(", ", info.RequiredParams)}" +
                (info.OptionalParams.Length > 0 ? $". Optional: {string.Join(", ", info.OptionalParams)}" : ""));

        return info.Builder(parameters);
    }

    // ── Builders ─────────────────────────────────────────────────────

    private static WebPartData BuildButton(Dictionary<string, string> p)
    {
        return new WebPartData
        {
            DataVersion = "1.0",
            Title = "Button",
            ServerProcessedContent = new ServerProcessedContent
            {
                SearchablePlainTexts = [new MetaDataKeyStringPair { Key = "button.label", Value = p["text"] }],
                Links = [new MetaDataKeyStringPair { Key = "button.linkUrl", Value = p["link"] }]
            },
            Properties = BuildUntypedProps(p.TryGetValue("alignment", out var a)
                ? new() { ["alignment"] = a } : [])
        };
    }

    private static WebPartData BuildDivider(Dictionary<string, string> _)
    {
        return new WebPartData
        {
            DataVersion = "1.0",
            Title = "Divider",
            Properties = BuildUntypedProps([])
        };
    }

    private static WebPartData BuildSpacer(Dictionary<string, string> p)
    {
        var height = p.TryGetValue("height", out var h) ? h : "40";
        return new WebPartData
        {
            DataVersion = "1.0",
            Title = "Spacer",
            Properties = BuildUntypedProps(new() { ["height"] = height })
        };
    }

    private static WebPartData BuildImage(Dictionary<string, string> p)
    {
        var props = new Dictionary<string, string> { ["imageSourceType"] = "2" };
        if (p.TryGetValue("altText", out var alt)) props["altText"] = alt;
        if (p.TryGetValue("caption", out var cap)) props["captionText"] = cap;

        return new WebPartData
        {
            DataVersion = "1.9",
            Title = "Image",
            Description = "Show an image on your page",
            Properties = BuildUntypedProps(props),
            ServerProcessedContent = new ServerProcessedContent
            {
                ImageSources = [new MetaDataKeyStringPair { Key = "imageSource", Value = p["imageUrl"] }]
            }
        };
    }

    private static WebPartData BuildLinkPreview(Dictionary<string, string> p)
    {
        return new WebPartData
        {
            DataVersion = "1.0",
            Title = "Link Preview",
            ServerProcessedContent = new ServerProcessedContent
            {
                Links = [new MetaDataKeyStringPair { Key = "link", Value = p["link"] }]
            },
            Properties = BuildUntypedProps([])
        };
    }

    // ── Helpers ──────────────────────────────────────────────────────

    private static UntypedNode BuildUntypedProps(Dictionary<string, string> props)
    {
        var dict = props.ToDictionary(
            kvp => kvp.Key,
            kvp => (UntypedNode)new UntypedString(kvp.Value));
        return new UntypedObject(dict);
    }
}
