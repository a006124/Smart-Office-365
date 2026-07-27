using Microsoft.Graph;
using Microsoft.Graph.Models;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Services
{
    public class SharePointSelectionService : ISharePointSelectionService
    {
        private readonly IGraphAuthService _authService;
        private string? _activeSiteId;

        public SharePointSelectionService(IGraphAuthService authService)
        {
            _authService = authService;
        }

        public async Task<List<SharePointSiteInfo>> GetAvailableSitesAsync()
        {
            var client = await _authService.GetAuthenticatedClientAsync();
            var sites = await client.Sites.GetAsync(
                config => config.QueryParameters.Filter = "siteCollection ne null"
            );

            var result = new List<SharePointSiteInfo>();
            if (sites?.Value != null)
            {
                foreach (var site in sites.Value)
                {
                    result.Add(new SharePointSiteInfo
                    {
                        Id = site.Id ?? string.Empty,
                        Name = site.Name ?? string.Empty,
                        DisplayName = site.DisplayName ?? site.Name ?? string.Empty,
                        WebUrl = site.WebUrl ?? string.Empty
                    });
                }
            }

            return result;
        }

        public Task SetActiveSiteAsync(string siteId)
        {
            _activeSiteId = siteId;
            return Task.CompletedTask;
        }

        public string GetActiveSiteId()
        {
            return _activeSiteId ?? string.Empty;
        }

        public bool HasActiveSite()
        {
            return !string.IsNullOrEmpty(_activeSiteId);
        }
    }
}