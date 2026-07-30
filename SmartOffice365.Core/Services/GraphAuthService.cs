using System;
using System.Net.Http;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using SmartOffice365.Core.Interfaces;

namespace SmartOffice365.Core.Services
{
    public class GraphAuthService : IGraphAuthService
    {
        private GraphServiceClient? _graphClient;
        private string _userDisplayName = string.Empty;

        // URL par défaut (Renault)
        private string _sharePointUrl = "https://grouperenault.sharepoint.com/sites/ShutdownMaintenance/";
        private string? _tenantId;

        private static readonly string[] Scopes = new[]
        {
            "https://graph.microsoft.com/.default"
        };

        /// <summary>
        /// Permet de mettre à jour l'URL cible et force la reconnexion si nécessaire
        /// </summary>
        public void UpdateSharePointUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url) || url.Equals(_sharePointUrl, StringComparison.OrdinalIgnoreCase))
                return;

            _sharePointUrl = url;
            _graphClient = null; // Force la ré-authentification avec le nouveau Tenant
            _tenantId = null;
        }

        public async Task<GraphServiceClient> GetAuthenticatedClientAsync()
        {
            if (_graphClient != null)
                return _graphClient;

            // 1. Résolution dynamique du Tenant ID à partir de l'URL SharePoint
            if (string.IsNullOrEmpty(_tenantId))
            {
                _tenantId = await ResolveTenantIdFromUrlAsync(_sharePointUrl);
            }

            // 2. Configuration de DefaultAzureCredential avec le Tenant ID dynamique
            var credential = new DefaultAzureCredential(
                new DefaultAzureCredentialOptions
                {
                    TenantId = _tenantId,
                    ExcludeSharedTokenCacheCredential = false,
                    ExcludeVisualStudioCredential = false,
                    ExcludeInteractiveBrowserCredential = false,
                    ExcludeVisualStudioCodeCredential = true,
                    ExcludeAzureCliCredential = true,
                    ExcludeAzurePowerShellCredential = true,
                    ExcludeManagedIdentityCredential = true,
                    ExcludeEnvironmentCredential = true
                });

            _graphClient = new GraphServiceClient(credential, Scopes);
            return _graphClient;
        }

        /// <summary>
        /// Interroge l'API Microsoft pour trouver le Tenant ID associé au domaine de l'URL
        /// </summary>
        private async Task<string> ResolveTenantIdFromUrlAsync(string url)
        {
            try
            {
                var uri = new Uri(url);
                string domain = uri.Host; // ex: grouperenault.sharepoint.com

                using var httpClient = new HttpClient();
                string configUrl = $"https://login.microsoftonline.com/{domain}/.well-known/openid-configuration";

                var response = await httpClient.GetStringAsync(configUrl);

                // On cherche l'ID (GUID) dans les URLs de la réponse (ex: login.microsoftonline.com/{GUID}/...)
                int index = response.IndexOf("login.microsoftonline.com/");
                if (index != -1)
                {
                    string sub = response.Substring(index + "login.microsoftonline.com/".Length);
                    int slashIndex = sub.IndexOf("/");
                    if (slashIndex != -1)
                    {
                        string tenantId = sub.Substring(0, slashIndex);
                        if (Guid.TryParse(tenantId, out _))
                        {
                            System.Diagnostics.Debug.WriteLine($"[GraphAuthService] Tenant ID résolu dynamiquement : {tenantId}");
                            return tenantId;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GraphAuthService] Échec de résolution du Tenant ID : {ex.Message}");
            }

            // Fallback de secours (Renault) si la résolution échoue
            return "d6b0bbee-7cd9-4d60-bce6-4a67b543e2ae";
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
                _userDisplayName = !string.IsNullOrWhiteSpace(user?.DisplayName)
                    ? user.DisplayName
                    : "Compte Office 365";
                return _userDisplayName;
            }
            catch
            {
                return Environment.UserName ?? "Utilisateur Renault";
            }
        }
    }
}
