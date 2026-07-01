using System.CommandLine;
using GraphCli.Services;

namespace GraphCli.Commands;

public static class AuthCommands
{
    public static Command Build(Option<string> formatOption)
    {
        var authCommand = new Command("auth", "Authentication management");
        authCommand.Subcommands.Add(BuildLogin(formatOption));
        authCommand.Subcommands.Add(BuildStatus(formatOption));
        authCommand.Subcommands.Add(BuildList(formatOption));
        authCommand.Subcommands.Add(BuildLogout(formatOption));
        return authCommand;
    }

    private static Command BuildLogin(Option<string> formatOption)
    {
        var tenantOption = new Option<string?>("--tenant")
        {
            Description = "Entra tenant ID. Required for a new profile (reused if the profile already exists)."
        };
        var clientOption = new Option<string?>("--client")
        {
            Description = "App (client) ID. Required for a new profile (reused if the profile already exists)."
        };

        var loginCommand = new Command("login", "Authenticate a profile to Microsoft Graph (select it with the global --profile option)")
        {
            tenantOption, clientOption
        };
        loginCommand.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var tenant = parseResult.GetValue(tenantOption);
                var client = parseResult.GetValue(clientOption);
                var format = parseResult.GetValue(formatOption) ?? "json";

                var auth = new AuthService();
                var result = await auth.LoginAsync(tenant, client);
                OutputService.Print(new
                {
                    status = "success",
                    profile = AuthService.Profile,
                    username = result.Account.Username,
                    expiresOn = result.ExpiresOn.ToString("o")
                }, format);
            }
            catch (Exception ex)
            {
                OutputService.PrintError("auth_failed", ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return loginCommand;
    }

    private static Command BuildStatus(Option<string> formatOption)
    {
        var statusCommand = new Command("status", "Show login status of the selected profile (choose it with --profile)");
        statusCommand.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var format = parseResult.GetValue(formatOption) ?? "json";
                var auth = new AuthService();
                var status = await auth.GetStatusAsync();
                OutputService.Print(status, format);
            }
            catch (Exception ex)
            {
                OutputService.PrintError("status_failed", ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return statusCommand;
    }

    private static Command BuildList(Option<string> formatOption)
    {
        var listCommand = new Command("list", "List all authenticated profiles");
        listCommand.SetAction((parseResult, ct) =>
        {
            var format = parseResult.GetValue(formatOption) ?? "json";
            var accounts = AuthService.GetAllAccounts();

            if (accounts.Count == 0)
            {
                OutputService.Print(new { message = "No profiles found. Run 'graph-cli auth login'." }, format);
                return Task.CompletedTask;
            }

            var list = accounts.Select(kvp => new
            {
                profile = kvp.Key,
                tenantId = kvp.Value.TenantId,
                username = kvp.Value.Username,
                selected = kvp.Key == AuthService.Profile
            }).ToArray();

            OutputService.Print(list, format);
            return Task.CompletedTask;
        });
        return listCommand;
    }

    private static Command BuildLogout(Option<string> formatOption)
    {
        var profileArg = new Argument<string?>("profile")
        {
            Description = "Profile to log out. Omit to log out the selected profile (--profile, else 'default').",
            Arity = ArgumentArity.ZeroOrOne
        };
        var logoutCommand = new Command("logout", "Remove a profile's cached credentials") { profileArg };
        logoutCommand.SetAction(async (parseResult, ct) =>
        {
            try
            {
                var profile = parseResult.GetValue(profileArg);
                var format = parseResult.GetValue(formatOption) ?? "json";

                var auth = new AuthService();
                var removed = await auth.LogoutAsync(profile);
                if (removed != null)
                {
                    var accounts = AuthService.GetAllAccounts();
                    OutputService.Print(new
                    {
                        status = "logged_out",
                        removedProfile = removed,
                        remainingProfiles = accounts.Count
                    }, format);
                }
                else
                {
                    OutputService.PrintError("not_logged_in", "No matching profile found.");
                    Environment.ExitCode = 1;
                }
            }
            catch (Exception ex)
            {
                OutputService.PrintError("logout_failed", ex.Message);
                Environment.ExitCode = 1;
            }
        });
        return logoutCommand;
    }
}
