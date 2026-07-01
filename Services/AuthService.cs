using System.Text.Json;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace GraphCli.Services;

public class AuthService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".graph-cli");
    private static readonly string ConfigPath = Path.Combine(ConfigDir, "config.json");

    private static readonly JsonSerializerOptions WriteOptions = new() { WriteIndented = true };

    private static readonly string[] DefaultScopes =
    [
        "User.Read", "User.ReadBasic.All",
        "Mail.ReadWrite", "Mail.Send",
        "Calendars.Read.Shared", "Calendars.ReadWrite",
        "Chat.Create", "Chat.ReadWrite", "ChatMessage.Read", "ChatMessage.Send",
        "Presence.Read.All",
        "Tasks.ReadWrite",
        "Files.Read.All", "Sites.ReadWrite.All"
    ];

    private IPublicClientApplication? _pca;

    // ----------------------------------------------------------------------
    // Profile — the single account this process is pinned to, resolved ONCE.
    //   precedence: --profile flag (set in Program.cs) > GRAPH_CLI_PROFILE env > "default".
    // There is no persisted "active" pointer: selection is per-process, never
    // mutated at runtime, so concurrent CLI/MCP processes never interfere.
    // ----------------------------------------------------------------------
    public static string Profile { get; set; } =
        Environment.GetEnvironmentVariable("GRAPH_CLI_PROFILE") is { Length: > 0 } p ? p : "default";

    // ----------------------------------------------------------------------
    // PublicClientApplication — built from the resolved profile's tenant/client.
    // The MSAL token cache (token-cache.bin) is shared across accounts; MSAL
    // keys entries by clientId + account, so multiple identities coexist in it.
    // ----------------------------------------------------------------------

    public async Task<IPublicClientApplication> GetPcaAsync()
    {
        if (_pca != null) return _pca;
        _pca = await BuildPcaAsync(ResolveConfig());
        return _pca;
    }

    private static async Task<IPublicClientApplication> BuildPcaAsync(GraphCliConfig config)
    {
        var pca = PublicClientApplicationBuilder
            .Create(config.ClientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, config.TenantId)
            .WithRedirectUri("http://localhost")
            .Build();

        await RegisterCacheAsync(pca);
        return pca;
    }

    /// <summary>
    /// Selects the MSAL cached account for this profile. Pins on HomeAccountId when
    /// captured (deterministic even when several profiles share one clientId). If the
    /// profile was never pinned (e.g. migrated legacy config), only a single-identity
    /// cache is unambiguous; otherwise we refuse rather than guess with FirstOrDefault.
    /// </summary>
    private static async Task<IAccount?> GetMsalAccountAsync(IPublicClientApplication pca, GraphCliConfig config)
    {
        var accounts = (await pca.GetAccountsAsync()).ToList();
        if (!string.IsNullOrEmpty(config.HomeAccountId))
            return accounts.FirstOrDefault(a => a.HomeAccountId?.Identifier == config.HomeAccountId);

        if (accounts.Count <= 1)
            return accounts.FirstOrDefault();

        throw new GraphCliConfigException(
            $"Profile '{Profile}' isn't pinned to an identity and the token cache holds " +
            $"{accounts.Count} accounts. Run 'graph-cli auth login" +
            (Profile == "default" ? "" : $" --profile {Profile}") + "' to pin the right one.");
    }

    // ----------------------------------------------------------------------
    // Token acquisition (active account)
    // ----------------------------------------------------------------------

    public async Task<string> GetAccessTokenAsync()
    {
        ConsentService.EnsureConsented();
        var config = ResolveConfig();
        var pca = await GetPcaAsync();
        var account = await GetMsalAccountAsync(pca, config);

        try
        {
            var result = await pca.AcquireTokenSilent(config.Scopes, account).ExecuteAsync();
            return result.AccessToken;
        }
        catch (MsalUiRequiredException)
        {
            var result = await pca.AcquireTokenInteractive(config.Scopes)
                .WithUseEmbeddedWebView(false)
                .ExecuteAsync();
            PersistAccountIdentity(config, result.Account);
            return result.AccessToken;
        }
    }

    // ----------------------------------------------------------------------
    // Login / switch / logout / status
    // ----------------------------------------------------------------------

    /// <summary>
    /// Authenticates the resolved <see cref="Profile"/> and persists it. Tenant/client come
    /// from the (optional) arguments, the existing profile of the same name, or the
    /// GRAPH_CLI_* environment variables. Scopes are always the current defaults.
    /// </summary>
    public async Task<AuthenticationResult> LoginAsync(string? tenantId, string? clientId)
    {
        ConsentService.EnsureConsented();
        var store = LoadStore();

        if (!store.Accounts.TryGetValue(Profile, out var config))
            config = new GraphCliConfig();

        var tenant = !string.IsNullOrEmpty(tenantId) ? tenantId
            : !string.IsNullOrEmpty(config.TenantId) ? config.TenantId
            : Environment.GetEnvironmentVariable("GRAPH_CLI_TENANT_ID");
        var client = !string.IsNullOrEmpty(clientId) ? clientId
            : !string.IsNullOrEmpty(config.ClientId) ? config.ClientId
            : Environment.GetEnvironmentVariable("GRAPH_CLI_CLIENT_ID");

        if (string.IsNullOrEmpty(tenant) || string.IsNullOrEmpty(client))
            throw new GraphCliConfigException(
                $"Profile '{Profile}' has no tenant/client. Pass --tenant <id> --client <id> " +
                "(or set GRAPH_CLI_TENANT_ID and GRAPH_CLI_CLIENT_ID).");

        config.TenantId = tenant;
        config.ClientId = client;
        config.Scopes = DefaultScopes;

        var pca = await BuildPcaAsync(config);
        var result = await pca.AcquireTokenInteractive(config.Scopes)
            .WithUseEmbeddedWebView(false)
            .ExecuteAsync();

        config.Username = result.Account.Username;
        config.HomeAccountId = result.Account.HomeAccountId?.Identifier;

        store.Accounts[Profile] = config;
        SaveStore(store);
        return result;
    }

    /// <summary>Returns all configured profiles.</summary>
    public static Dictionary<string, GraphCliConfig> GetAllAccounts() => LoadStore().Accounts;

    /// <summary>
    /// Removes one profile (the resolved <see cref="Profile"/> if not specified),
    /// evicting its MSAL cache entry. Returns the removed profile name, or null if
    /// nothing matched.
    /// </summary>
    public async Task<string?> LogoutAsync(string? profile = null)
    {
        var store = LoadStore();
        if (store.Accounts.Count == 0)
            return null;

        var key = profile ?? Profile;
        if (!store.Accounts.TryGetValue(key, out var config))
            return null;

        try
        {
            var pca = await BuildPcaAsync(config);
            var accounts = await pca.GetAccountsAsync();
            // Remove only this profile's identity when we can pin it; otherwise all
            // accounts under this clientId (legacy profiles with no captured id).
            var targets = !string.IsNullOrEmpty(config.HomeAccountId)
                ? accounts.Where(a => a.HomeAccountId?.Identifier == config.HomeAccountId).ToList()
                : accounts.ToList();
            foreach (var account in targets)
                await pca.RemoveAsync(account);
        }
        catch { /* best-effort cache eviction */ }

        store.Accounts.Remove(key);

        if (store.Accounts.Count == 0)
        {
            if (File.Exists(ConfigPath))
                File.Delete(ConfigPath);
            return key;
        }

        SaveStore(store);
        return key;
    }

    /// <summary>Login status of the resolved <see cref="Profile"/> (used by the MCP auth_status tool).</summary>
    public async Task<AuthStatus> GetStatusAsync()
    {
        var store = LoadStore();
        if (!store.Accounts.TryGetValue(Profile, out var config))
            return new AuthStatus
            {
                IsLoggedIn = false,
                Profile = Profile,
                Message = $"Profile '{Profile}' not found. Run 'graph-cli auth login" +
                          (Profile == "default" ? "" : $" --profile {Profile}") + "'."
            };

        config.Scopes = DefaultScopes;
        var pca = await BuildPcaAsync(config);
        var account = await GetMsalAccountAsync(pca, config);

        if (account == null)
            return new AuthStatus
            {
                IsLoggedIn = false,
                Profile = Profile,
                Username = config.Username,
                Message = "No cached token. Run 'graph-cli auth login' to authenticate."
            };

        try
        {
            var result = await pca.AcquireTokenSilent(config.Scopes, account).ExecuteAsync();
            return new AuthStatus
            {
                IsLoggedIn = true,
                Profile = Profile,
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
                Profile = Profile,
                Username = account.Username,
                Message = "Token expired. Run 'graph-cli auth login' to re-authenticate."
            };
        }
    }

    private static void PersistAccountIdentity(GraphCliConfig config, IAccount account)
    {
        var store = LoadStore();
        if (!store.Accounts.TryGetValue(Profile, out var stored))
            return;
        stored.Username = account.Username;
        stored.HomeAccountId = account.HomeAccountId?.Identifier;
        SaveStore(store);
    }

    // ----------------------------------------------------------------------
    // Config store (multi-account) with legacy single-account migration
    // ----------------------------------------------------------------------

    private static GraphCliConfig ResolveConfig()
    {
        var store = LoadStore();
        if (store.Accounts.TryGetValue(Profile, out var config))
        {
            // Always use DefaultScopes so new permissions are picked up on upgrade
            config.Scopes = DefaultScopes;
            return config;
        }

        var envTenant = Environment.GetEnvironmentVariable("GRAPH_CLI_TENANT_ID");
        var envClient = Environment.GetEnvironmentVariable("GRAPH_CLI_CLIENT_ID");
        if (!string.IsNullOrEmpty(envTenant) && !string.IsNullOrEmpty(envClient))
        {
            return new GraphCliConfig
            {
                TenantId = envTenant,
                ClientId = envClient,
                Scopes = DefaultScopes
            };
        }

        throw new GraphCliConfigException(
            $"Profile '{Profile}' not configured. Run:\n" +
            "  graph-cli auth login" + (Profile == "default" ? "" : $" --profile {Profile}") +
            " --tenant <tenant-id> --client <client-id>\n" +
            "or set GRAPH_CLI_TENANT_ID and GRAPH_CLI_CLIENT_ID environment variables.");
    }

    private static ConfigStore LoadStore()
    {
        if (!File.Exists(ConfigPath))
            return new ConfigStore();

        var json = File.ReadAllText(ConfigPath);

        // New multi-account format
        try
        {
            var store = JsonSerializer.Deserialize<ConfigStore>(json);
            if (store?.Accounts.Count > 0)
                return store;
        }
        catch { }

        // Legacy single-account format { tenantId, clientId, scopes } — migrate to "default"
        try
        {
            var legacy = JsonSerializer.Deserialize<GraphCliConfig>(json);
            if (legacy != null && !string.IsNullOrEmpty(legacy.TenantId) && !string.IsNullOrEmpty(legacy.ClientId))
            {
                var store = new ConfigStore();
                store.Accounts["default"] = legacy;
                SaveStore(store);
                return store;
            }
        }
        catch { }

        return new ConfigStore();
    }

    private static void SaveStore(ConfigStore store)
    {
        Directory.CreateDirectory(ConfigDir);
        File.WriteAllText(ConfigPath, JsonSerializer.Serialize(store, WriteOptions));
    }

    private static async Task RegisterCacheAsync(IPublicClientApplication pca)
    {
        Directory.CreateDirectory(ConfigDir);
        var storageProperties = new StorageCreationPropertiesBuilder("token-cache.bin", ConfigDir)
            .WithLinuxKeyring(
                schemaName: "com.graphcli.tokencache",
                collection: "default",
                secretLabel: "graph-cli MSAL token cache",
                attribute1: new KeyValuePair<string, string>("Version", "1"),
                attribute2: new KeyValuePair<string, string>("ProductGroup", "graph-cli"))
            .WithMacKeyChain(
                serviceName: "graph-cli",
                accountName: "graph-cli-msal-cache")
            .Build();

        var cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
        cacheHelper.RegisterCache(pca.UserTokenCache);
    }
}

public class ConfigStore
{
    public Dictionary<string, GraphCliConfig> Accounts { get; set; } = new();
}

public class GraphCliConfig
{
    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string[] Scopes { get; set; } = [];

    /// <summary>Username captured from MSAL after the last interactive login (for display + account selection).</summary>
    public string? Username { get; set; }

    /// <summary>MSAL HomeAccountId.Identifier captured at login; pins which cached identity this profile uses.</summary>
    public string? HomeAccountId { get; set; }
}

public class AuthStatus
{
    public bool IsLoggedIn { get; set; }
    public string? Profile { get; set; }
    public string? Username { get; set; }
    public string? Environment { get; set; }
    public DateTimeOffset? ExpiresOn { get; set; }
    public string? Message { get; set; }
}

public class GraphCliConfigException : Exception
{
    public GraphCliConfigException(string message) : base(message) { }
}
