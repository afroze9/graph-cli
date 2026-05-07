using System.CommandLine;
using GraphCli.Services;

namespace GraphCli.Commands;

public static class ConsentCommands
{
    public static Command Build()
    {
        var consentCommand = new Command("consent", "Manage consent to graph-cli terms of use");

        var showCommand = new Command("show", "Show the consent terms");
        showCommand.SetAction((_, _) =>
        {
            Console.WriteLine(ConsentService.ConsentText);
            return Task.FromResult(0);
        });

        var grantCommand = new Command("grant", "Accept the terms and grant consent to use graph-cli");
        grantCommand.SetAction((_, _) =>
        {
            try
            {
                Console.WriteLine(ConsentService.ConsentText);
                Console.WriteLine();
                var status = ConsentService.Grant();
                OutputService.Print(new
                {
                    status = "consent_granted",
                    consentedAt = status.ConsentedAt?.ToString("o"),
                    version = status.Version
                });
                return Task.FromResult(0);
            }
            catch (Exception ex)
            {
                OutputService.PrintError("consent_grant_failed", ex.Message);
                Environment.ExitCode = 1;
                return Task.FromResult(1);
            }
        });

        var statusCommand = new Command("status", "Show current consent status");
        statusCommand.SetAction((_, _) =>
        {
            try
            {
                OutputService.Print(ConsentService.GetStatus());
                return Task.FromResult(0);
            }
            catch (Exception ex)
            {
                OutputService.PrintError("consent_status_failed", ex.Message);
                Environment.ExitCode = 1;
                return Task.FromResult(1);
            }
        });

        var revokeCommand = new Command("revoke", "Revoke previously granted consent");
        revokeCommand.SetAction((_, _) =>
        {
            try
            {
                ConsentService.Revoke();
                OutputService.Print(new { status = "consent_revoked" });
                return Task.FromResult(0);
            }
            catch (Exception ex)
            {
                OutputService.PrintError("consent_revoke_failed", ex.Message);
                Environment.ExitCode = 1;
                return Task.FromResult(1);
            }
        });

        consentCommand.Subcommands.Add(showCommand);
        consentCommand.Subcommands.Add(grantCommand);
        consentCommand.Subcommands.Add(statusCommand);
        consentCommand.Subcommands.Add(revokeCommand);
        return consentCommand;
    }
}
