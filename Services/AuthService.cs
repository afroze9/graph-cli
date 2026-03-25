using System.Text.Json;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace GraphCli.Services;

public class AuthService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".graph-cli");
    private static readonly string ConfigPath = Path.Combine(ConfigDir, "config.json");

    private static readonly string[] DefaultScopes =
    [
        "User.Read", "User.ReadBasic.All",
        "Mail.ReadWrite", "Mail.Send",
        "Calendars.Read.Shared", "Calendars.ReadWrite",
        "Chat.Create", "Chat.ReadWrite", "ChatMessage.Read", "ChatMessage.Send",
        "Presence.Read.All",
        "Tasks.ReadWrite"
    ];

    private IPublicClientApplication? _pca;

    public async Task<IPublicClientApplication> GetPcaAsync()
    {
        if (_pca != null) return _pca;

        var config = LoadOrCreateConfig();
        _pca = PublicClientApplicationBuilder
            .Create(config.ClientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, config.TenantId)
            .WithRedirectUri("http://localhost")
            .Build();

        await RegisterCacheAsync(_pca);
        return _pca;
    }

    public async Task<string> GetAccessTokenAsync()
    {
        var pca = await GetPcaAsync();
        var accounts = await pca.GetAccountsAsync();
        var config = LoadOrCreateConfig();

        try
        {
            var result = await pca.AcquireTokenSilent(config.Scopes, accounts.FirstOrDefault())
                .ExecuteAsync();
            return result.AccessToken;
        }
        catch (MsalUiRequiredException)
        {
            var result = await pca.AcquireTokenInteractive(config.Scopes)
                .WithUseEmbeddedWebView(false)
                .ExecuteAsync();
            return result.AccessToken;
        }
    }

    public async Task<AuthenticationResult> LoginAsync()
    {
        var pca = await GetPcaAsync();
        var config = LoadOrCreateConfig();
        return await pca.AcquireTokenInteractive(config.Scopes)
            .WithUseEmbeddedWebView(false)
            .ExecuteAsync();
    }

    public async Task LogoutAsync()
    {
        var pca = await GetPcaAsync();
        var accounts = await pca.GetAccountsAsync();
        foreach (var account in accounts)
            await pca.RemoveAsync(account);
    }

    public async Task<AuthStatus> GetStatusAsync()
    {
        var pca = await GetPcaAsync();
        var accounts = await pca.GetAccountsAsync();
        var account = accounts.FirstOrDefault();

        if (account == null)
            return new AuthStatus { IsLoggedIn = false };

        var config = LoadOrCreateConfig();
        try
        {
            var result = await pca.AcquireTokenSilent(config.Scopes, account)
                .ExecuteAsync();
            return new AuthStatus
            {
                IsLoggedIn = true,
                Username = account.Username,
                Environment = account.Environment,
                ExpiresOn = result.ExpiresOn
            };
        }
        catch (MsalUiRequiredException)
        {
            return new AuthStatus
            {
                IsLoggedIn = false,
                Username = account.Username,
                Message = "Token expired. Run 'graph-cli auth login' to re-authenticate."
            };
        }
    }

    private static async Task RegisterCacheAsync(IPublicClientApplication pca)
    {
        Directory.CreateDirectory(ConfigDir);
        var storageProperties = new StorageCreationPropertiesBuilder("token-cache.bin", ConfigDir)
            .WithUnprotectedFile()
            .Build();

        var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
        cacheHelper.RegisterCache(pca.UserTokenCache);
    }

    private static GraphCliConfig LoadOrCreateConfig()
    {
        if (File.Exists(ConfigPath))
        {
            var json = File.ReadAllText(ConfigPath);
            var config = JsonSerializer.Deserialize<GraphCliConfig>(json);
            if (config != null) return config;
        }

        var envTenant = Environment.GetEnvironmentVariable("GRAPH_CLI_TENANT_ID");
        var envClient = Environment.GetEnvironmentVariable("GRAPH_CLI_CLIENT_ID");

        if (string.IsNullOrEmpty(envTenant) || string.IsNullOrEmpty(envClient))
        {
            Console.Error.WriteLine("No configuration found. Please create ~/.graph-cli/config.json with:");
            Console.Error.WriteLine("""
            {
              "tenantId": "<your-tenant-id>",
              "clientId": "<your-client-id>"
            }
            """);
            Console.Error.WriteLine("Or set GRAPH_CLI_TENANT_ID and GRAPH_CLI_CLIENT_ID environment variables.");
            Environment.Exit(1);
        }

        var newConfig = new GraphCliConfig
        {
            TenantId = envTenant!,
            ClientId = envClient!,
            Scopes = DefaultScopes
        };

        Directory.CreateDirectory(ConfigDir);
        File.WriteAllText(ConfigPath, JsonSerializer.Serialize(newConfig, new JsonSerializerOptions { WriteIndented = true }));
        return newConfig;
    }
}

public class GraphCliConfig
{
    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string[] Scopes { get; set; } = [];
}

public class AuthStatus
{
    public bool IsLoggedIn { get; set; }
    public string? Username { get; set; }
    public string? Environment { get; set; }
    public DateTimeOffset? ExpiresOn { get; set; }
    public string? Message { get; set; }
}
