using System.Net;
using System.Text.RegularExpressions;
using Microsoft.Graph;

namespace GraphCli.Services;

/// <summary>A single resolved @-mention: the address to notify and the name Outlook shows.</summary>
internal sealed record MentionTarget(string Address, string Name);

/// <summary>The rewritten HTML body plus the mentions to attach to the Graph request.</summary>
internal sealed record MentionPlan(string Body, List<MentionTarget> Targets);

// Turns the CLI's `--mentions` list plus `<at id="N">Name</at>` tags into an Outlook
// @-mention. The `<at>` syntax is deliberately the same one `graph-cli chat send` uses,
// so one convention covers both Teams and mail.
//
// Outlook needs two things for a real @-mention:
//   1. an anchor in the HTML body — <a href="mailto:addr">@Name</a> — which is what the
//      reader sees and what the person card hangs off;
//   2. the `mentions` collection on the message, which drives the @ glyph in the
//      recipient's message list and the "Mentioned mail" filter.
// The mentioned person must also be a recipient, or the mention notifies nobody.
internal static class MailMentionService
{
    public static async Task<MentionPlan> BuildAsync(
        GraphServiceClient client, string body, string contentType, string[] mentions)
    {
        if (!string.Equals(contentType, "html", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                "--mentions requires --content-type html. Reference each mention in the body with <at id=\"N\">Name</at> where N is the zero-based index matching --mentions.");

        var targets = new List<MentionTarget>();

        for (int i = 0; i < mentions.Length; i++)
        {
            var address = mentions[i].Trim();
            if (!address.Contains('@'))
                throw new ArgumentException(
                    $"--mentions: '{address}' is not an email address. Mail mentions must be email addresses, because the mentioned person is added as a recipient.");

            // The `<at>` tag carries the display text. Reading it back lets callers write
            // any text they like ("Afroze") without matching the full AAD displayName.
            var atMatch = Regex.Match(
                body,
                $"<at id=[\"']{i}[\"'][^>]*>(.*?)</at>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!atMatch.Success)
                throw new ArgumentException(
                    $"--mentions: body is missing <at id=\"{i}\">...</at> tag for mention #{i} ({address}). Add it to the --body HTML so Outlook can render the @-mention.");

            var mentionText = atMatch.Groups[1].Value.Trim();
            var name = await ResolveDisplayNameAsync(client, address) ?? mentionText;

            // Outlook writes mentions as "@Name" — add the prefix unless the caller did.
            var label = mentionText.StartsWith('@') ? mentionText : "@" + mentionText;
            var anchor = BuildAnchor(address, label);

            body = body.Remove(atMatch.Index, atMatch.Length).Insert(atMatch.Index, anchor);
            targets.Add(new MentionTarget(address, name));
        }

        return new MentionPlan(body, targets);
    }

    // Match the anchor Outlook itself writes for an @-mention:
    //   <a id="OWAAM<32 hex>" href="mailto:addr"><span style="text-decoration:none">@Name</span></a>
    //
    // The `id` is load-bearing, and its exact form matters. Tested against real mail with
    // five id variants in one message, all pointing at the same address, all backed by the
    // same valid `mentions` record. Only "OWAAM" + 32 uppercase hex rendered as a mention:
    //
    //   OWAAM + 32 hex   -> mention          (this is what Outlook Web writes)
    //   OWA   + 32 hex   -> plain blue link
    //   32 hex, no prefix-> plain blue link
    //   arbitrary string -> plain blue link
    //   no id at all     -> plain blue link
    //
    // So the `mentions` collection alone does not style the body. It drives the @ glyph and
    // the "Mentioned mail" filter, which are the parts Microsoft documents; the anchor id is
    // an undocumented Outlook detail. Keep the two concerns separate when this breaks: the
    // mention still works without the markup, it just looks like a link.
    //
    // The label text does not matter — a short "@Ali" renders the same as the full display
    // name. The span only suppresses the underline. Outlook writes one, so we do too, but it
    // is not what makes the mention render.
    private static string BuildAnchor(string address, string label)
    {
        var anchorId = "OWAAM" + Guid.NewGuid().ToString("N").ToUpperInvariant();
        return $"<a id=\"{anchorId}\" href=\"mailto:{WebUtility.HtmlEncode(address)}\">" +
               $"<span style=\"text-decoration:none\">{label}</span></a>";
    }

    // Best effort: a nicer display name when the address is in the directory. External
    // addresses are legitimate mention targets, so a failed lookup is not an error.
    private static async Task<string?> ResolveDisplayNameAsync(GraphServiceClient client, string address)
    {
        try
        {
            var user = await client.Users[address].GetAsync(r =>
            {
                r.QueryParameters.Select = ["id", "displayName"];
            });
            return user?.DisplayName;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError)
        {
            return null;
        }
    }
}
