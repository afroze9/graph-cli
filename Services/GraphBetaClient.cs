using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Graph.Models.ODataErrors;

namespace GraphCli.Services;

// Thin POST helper for the Microsoft Graph /beta endpoint.
//
// The generated v1.0 SDK has no model for the `mention` resource — Outlook @-mentions
// live only in /beta (see https://learn.microsoft.com/graph/api/resources/mention).
// Mention-bearing mail requests are therefore serialized to raw JSON and posted here.
// Failures are rethrown as ODataError so callers keep the same catch blocks they use
// for ordinary SDK calls.
internal static class GraphBetaClient
{
    private const string BaseUrl = "https://graph.microsoft.com/beta";
    private static readonly HttpClient Http = new();

    /// <summary>
    /// POSTs <paramref name="payload"/> to a /beta path (e.g. "/me/sendMail").
    /// Returns the parsed response body, or null when Graph answers 202/204 with no content.
    /// </summary>
    public static async Task<JsonObject?> PostAsync(string path, JsonNode payload)
    {
        var token = await new AuthService().GetAccessTokenAsync();

        using var request = new HttpRequestMessage(HttpMethod.Post, BaseUrl + path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Content = new StringContent(payload.ToJsonString(), Encoding.UTF8, "application/json");

        using var response = await Http.SendAsync(request);
        var text = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
            throw BuildError(response.StatusCode, text);

        return string.IsNullOrWhiteSpace(text) ? null : JsonNode.Parse(text) as JsonObject;
    }

    /// <summary>GETs a /beta path (e.g. "/me/messages/{id}?$expand=mentions").</summary>
    public static async Task<JsonObject?> GetAsync(string path)
    {
        var token = await new AuthService().GetAccessTokenAsync();

        using var request = new HttpRequestMessage(HttpMethod.Get, BaseUrl + path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        using var response = await Http.SendAsync(request);
        var text = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
            throw BuildError(response.StatusCode, text);

        return string.IsNullOrWhiteSpace(text) ? null : JsonNode.Parse(text) as JsonObject;
    }

    private static ODataError BuildError(HttpStatusCode status, string body)
    {
        var code = $"http_{(int)status}";
        var message = string.IsNullOrWhiteSpace(body) ? status.ToString() : body;

        try
        {
            if (JsonNode.Parse(body)?["error"] is JsonNode error)
            {
                code = error["code"]?.GetValue<string>() ?? code;
                message = error["message"]?.GetValue<string>() ?? message;
            }
        }
        catch (JsonException)
        {
            // Non-JSON error body (proxy/gateway page) — keep the raw text as the message.
        }

        return new ODataError { Error = new MainError { Code = code, Message = message } };
    }
}
