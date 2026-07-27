using Azure.Identity;
using Microsoft.Graph;
using SmartOffice365.Core.Interfaces;

namespace SmartOffice365.Core.Services
{
    public class GraphAuthService : IGraphAuthService
    {
        private GraphServiceClient? _graphClient;
        private string _userDisplayName = string.Empty;

        private static readonly string[] Scopes = new[]
        {
            "https://graph.microsoft.com/.default"
        };

        public async Task<GraphServiceClient> GetAuthenticatedClientAsync()
        {
            if (_graphClient != null) return _graphClient;

            // Utilise DefaultAzureCredential qui fonctionne avec la session Office 365
            var credential = new DefaultAzureCredential(
                new DefaultAzureCredentialOptions
                {
                    ExcludeInteractiveBrowserCredential = true,
                    ExcludeVisualStudioCredential = true,
                    ExcludeVisualStudioCodeCredential = true,
                    ExcludeAzureCliCredential = true,
                    ExcludeAzurePowerShellCredential = true,
                    ExcludeManagedIdentityCredential = true,
                    ExcludeEnvironmentCredential = true
                });

            _graphClient = new GraphServiceClient(credential, Scopes);
            
            // Test rapide pour valider l'authentification
            try
            {
                var user = await _graphClient.Me.GetAsync();
                _userDisplayName = user?.DisplayName ?? "Compte Office 365";
            }
            catch
            {
                _userDisplayName = "Non connecté (Office 365)";
            }

            return _graphClient;
        }

        public async Task<bool> IsAuthenticatedAsync()
        {
            try
            {
                var client = await GetAuthenticatedClientAsync();
                await client.Me.GetAsync();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task<string> GetCurrentUserDisplayNameAsync()
        {
            if (!string.IsNullOrEmpty(_userDisplayName))
                return _userDisplayName;

            try
            {
                var client = await GetAuthenticatedClientAsync();
                var user = await client.Me.GetAsync();
                _userDisplayName = user?.DisplayName ?? "Compte Office 365";
                return _userDisplayName;
            }
            catch
            {
                return "Non connecté (Office 365)";
            }
        }
    }
}