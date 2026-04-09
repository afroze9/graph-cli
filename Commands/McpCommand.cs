using System.CommandLine;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol;

namespace GraphCli.Commands;

public static class McpCommand
{
    public static Command Build()
    {
        var cmd = new Command("mcp", "Start MCP (Model Context Protocol) server over stdio");
        cmd.SetAction(async (parseResult, ct) =>
        {
            var builder = Host.CreateApplicationBuilder();
            builder.Logging.ClearProviders();
            builder.Services
                .AddMcpServer()
                .WithStdioServerTransport()
                .WithToolsFromAssembly();

            var app = builder.Build();
            await app.RunAsync(ct);
        });
        return cmd;
    }
}
