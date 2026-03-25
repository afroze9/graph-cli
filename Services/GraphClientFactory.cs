using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;

namespace GraphCli.Services;

public static class GraphClientProvider
{
    public static async Task<GraphServiceClient> CreateAsync()
    {
        var authService = new AuthService();
        var tokenProvider = new MsalTokenProvider(authService);
        var authProvider = new BaseBearerTokenAuthenticationProvider(tokenProvider);
        // Ensure token cache is initialized
        await authService.GetPcaAsync();
        return new GraphServiceClient(authProvider);
    }
}

internal class MsalTokenProvider : IAccessTokenProvider
{
    private readonly AuthService _authService;

    public MsalTokenProvider(AuthService authService) => _authService = authService;

    public async Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default)
    {
        return await _authService.GetAccessTokenAsync();
    }

    public AllowedHostsValidator AllowedHostsValidator { get; } =
        new(["graph.microsoft.com"]);
}
