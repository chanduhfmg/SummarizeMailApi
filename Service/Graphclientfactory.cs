using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;

namespace SummarizeMail.Service
{
    /// <summary>
    /// Builds and holds the single app-only GraphServiceClient used by every
    /// mail/calendar method in ExchangeMailService. Register this as a
    /// Singleton in DI (see Program.cs snippet below) — a ClientSecretCredential
    /// + GraphServiceClient pair is thread-safe and reusable across requests,
    /// unlike the old per-call "new ExchangeService(...)" pattern.
    /// </summary>
    public interface IGraphClientFactory
    {
        GraphServiceClient GetClient();
        string ReminderMailbox { get; }
    }

    public class GraphClientFactory : IGraphClientFactory
    {
        private readonly GraphServiceClient _client;
        private readonly string _reminderMailbox;

        public GraphClientFactory(IConfiguration configuration)
        {
            var tenantId = configuration["Graph:TenantId"];
            var clientId = configuration["Graph:ClientId"];
            var clientSecret = configuration["Graph:ClientSecret"];
             _reminderMailbox = configuration["Graph:ReminderMailbox"];

            if (string.IsNullOrWhiteSpace(tenantId) ||
                string.IsNullOrWhiteSpace(clientId) ||
                string.IsNullOrWhiteSpace(clientSecret) ||
                    string.IsNullOrWhiteSpace(_reminderMailbox)) 
            {
                throw new InvalidOperationException(
                    "Graph:TenantId, Graph:ClientId and Graph:ClientSecret must all be set " +
                    "in appsettings.json (or user secrets / environment) before the app starts.");
            }

            // App-only auth: no signed-in user, the app authenticates as itself
            // and every call below addresses a mailbox explicitly via .Users[address].
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

            _client = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
        }

        public GraphServiceClient GetClient() => _client;
        public string ReminderMailbox => _reminderMailbox;
    }
}