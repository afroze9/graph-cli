using System.ComponentModel;
using GraphCli.Services;
using ModelContextProtocol.Server;

namespace GraphCli.McpTools;

[McpServerToolType]
public static class ConsentTools
{
    [McpServerTool(Name = "consent_status"), Description("Check whether the user has granted consent to graph-cli's terms of use. All other graph-cli tools require consent before they will run.")]
    public static string Status()
    {
        try
        {
            return McpGraphHelper.ToJson(ConsentService.GetStatus());
        }
        catch (Exception ex)
        {
            return McpGraphHelper.HandleException(ex);
        }
    }

    [McpServerTool(Name = "consent_show"), Description("Return the graph-cli consent terms that the user must accept before other tools will work.")]
    public static string Show()
    {
        return McpGraphHelper.ToJson(new { terms = ConsentService.ConsentText });
    }

    [McpServerTool(Name = "consent_grant"), Description("Record that the user has read and accepted the graph-cli terms of use. Call only after presenting the terms (consent_show) to the user and receiving explicit acceptance.")]
    public static string Grant()
    {
        try
        {
            var status = ConsentService.Grant();
            return McpGraphHelper.ToJson(new
            {
                status = "consent_granted",
                consentedAt = status.ConsentedAt?.ToString("o"),
                version = status.Version
            });
        }
        catch (Exception ex)
        {
            return McpGraphHelper.HandleException(ex);
        }
    }

    [McpServerTool(Name = "consent_revoke"), Description("Revoke previously granted consent. After revocation other graph-cli tools will refuse to run until consent is granted again.")]
    public static string Revoke()
    {
        try
        {
            ConsentService.Revoke();
            return McpGraphHelper.ToJson(new { status = "consent_revoked" });
        }
        catch (Exception ex)
        {
            return McpGraphHelper.HandleException(ex);
        }
    }
}
