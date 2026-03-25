using System.CommandLine;
using GraphCli.Services;

namespace GraphCli.Commands;

public static class AuthCommands
{
    public static Command Build()
    {
        var authCommand = new Command("auth", "Authentication management");

        var loginCommand = new Command("login", "Authenticate to Microsoft Graph");
        loginCommand.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var auth = new AuthService();
                var result = await auth.LoginAsync();
                OutputService.Print(new
                {
                    status = "success",
                    username = result.Account.Username,
                    expiresOn = result.ExpiresOn.ToString("o")
                });
            }
            catch (Exception ex)
            {
                OutputService.PrintError("auth_failed", ex.Message);
                Environment.ExitCode = 1;
            }
        });

        var statusCommand = new Command("status", "Show current authentication status");
        statusCommand.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var auth = new AuthService();
                var status = await auth.GetStatusAsync();
                OutputService.Print(status);
            }
            catch (Exception ex)
            {
                OutputService.PrintError("status_failed", ex.Message);
                Environment.ExitCode = 1;
            }
        });

        var logoutCommand = new Command("logout", "Clear cached tokens");
        logoutCommand.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var auth = new AuthService();
                await auth.LogoutAsync();
                OutputService.Print(new { status = "logged_out" });
            }
            catch (Exception ex)
            {
                OutputService.PrintError("logout_failed", ex.Message);
                Environment.ExitCode = 1;
            }
        });

        authCommand.Subcommands.Add(loginCommand);
        authCommand.Subcommands.Add(statusCommand);
        authCommand.Subcommands.Add(logoutCommand);
        return authCommand;
    }
}
