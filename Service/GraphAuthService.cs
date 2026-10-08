using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;

public class GraphAuthService
{
    private readonly IConfidentialClientApplication _msalClient;
    private readonly string[] _scopes;
    private readonly string _redirectUri;

    public GraphAuthService(IConfiguration config)
    {
        // Use a dedicated section for the delegated-login app registration —
        // keep this separate from Joseph's app-only "Graph" section for SummarizeMail.
        var azureAd = config.GetSection("Graph");
        _redirectUri = azureAd["RedirectUri"];
        _scopes = config.GetSection("GraphScopes").Get<string[]>();

        _msalClient = ConfidentialClientApplicationBuilder
            .Create(azureAd["ClientId"])
            .WithClientSecret(azureAd["ClientSecret"])
            .WithAuthority(new Uri($"{azureAd["Instance"]}{azureAd["TenantId"]}"), validateAuthority: true)
            .WithRedirectUri(_redirectUri)
            .Build();
    }

    public async Task<Uri> GetLoginUrlAsync()
    {
        return await _msalClient
            .GetAuthorizationRequestUrl(_scopes)
            .WithRedirectUri(_redirectUri)
            .ExecuteAsync();
    }

    public async Task<AuthenticationResult> CompleteLoginAsync(string code)
    {
        return await _msalClient
            .AcquireTokenByAuthorizationCode(_scopes, code)
            .ExecuteAsync();
    }

    public async Task<AuthenticationResult> GetTokenForUserAsync(string userIdentifier)
    {
        var accounts = await _msalClient.GetAccountsAsync();
        var account = accounts.FirstOrDefault(a => a.HomeAccountId.Identifier.Contains(userIdentifier));

        if (account == null)
            throw new InvalidOperationException("Session expired or not found — user needs to log in again.");

        return await _msalClient
            .AcquireTokenSilent(_scopes, account)
            .ExecuteAsync();
    }

    // Pulls display name + job title (the "role" to show in the UI) from Graph
    public async Task<(string DisplayName, string Email, string JobTitle)> GetUserProfileAsync(string accessToken)
    {
        var tokenProvider = new SimpleAccessTokenProvider(accessToken);
        var authProvider = new BaseBearerTokenAuthenticationProvider(tokenProvider);
        var graphClient = new GraphServiceClient(authProvider);

        var me = await graphClient.Me.GetAsync(requestConfig =>
        {
            requestConfig.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName", "jobTitle" };
        });

        return (me.DisplayName, me.Mail ?? me.UserPrincipalName, me.JobTitle);
    }
}

// Wraps an already-acquired MSAL access token so Graph's v5 client can use it directly
public class SimpleAccessTokenProvider : IAccessTokenProvider
{
    private readonly string _token;
    public SimpleAccessTokenProvider(string token) => _token = token;

    public Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object> additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default)
        => Task.FromResult(_token);

    public AllowedHostsValidator AllowedHostsValidator { get; } = new AllowedHostsValidator();
}