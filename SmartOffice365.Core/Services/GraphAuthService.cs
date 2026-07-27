using Azure.Identity;
using Microsoft.Graph;
using SmartOffice365.Core.Interfaces;

namespace SmartOffice365.Core.Services
{
    public class GraphAuthService : IGraphAuthService
    {
        private GraphServiceClient? _graphClient;
        private InteractiveBrowserCredential? _credential;
        private string _userDisplayName = string.Empty;

        private static readonly string[] Scopes = new[]
        {
            "User.Read",
            "Sites.ReadWrite.All",
            "ChannelMessage.Send",
            "Mail.Send"
        };

        public async Task<GraphServiceClient> GetAuthenticatedClientAsync()
        {
            if (_graphClient != null) return _graphClient;

            var options = new InteractiveBrowserCredentialOptions
            {
                TokenCachePersistenceOptions = new TokenCachePersistenceOptions()
            };

            _credential = new InteractiveBrowserCredential(options);
            _graphClient = new GraphServiceClient(_credential, Scopes);

            return _graphClient;
        }

        public async Task<bool> SignInAsync()
        {
            try
            {
                var client = await GetAuthenticatedClientAsync();
                var user = await client.Me.GetAsync();
                _userDisplayName = user?.DisplayName ?? "Compte Office 365";
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public async Task SignOutAsync()
        {
            _graphClient = null;
            _credential = null;
            _userDisplayName = string.Empty;
            await Task.CompletedTask;
        }

        public async Task<bool> IsAuthenticatedAsync()
        {
            return _graphClient != null;
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
