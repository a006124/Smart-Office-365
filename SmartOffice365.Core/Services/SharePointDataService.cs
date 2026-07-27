public class SharePointDataService : ISharePointDataService
{
    private readonly IGraphAuthService _authService;
    private readonly ISharePointSelectionService _selectionService;
    private string? _siteId;

    public SharePointDataService(
        IGraphAuthService authService, 
        ISharePointSelectionService selectionService)
    {
        _authService = authService;
        _selectionService = selectionService;
    }

    public void SetSiteId(string siteId)
    {
        _siteId = siteId;
    }

    private async Task<string> GetSiteIdAsync()
    {
        if (!string.IsNullOrEmpty(_siteId))
            return _siteId;

        if (_selectionService.HasActiveSite())
        {
            _siteId = _selectionService.GetActiveSiteId();
            return _siteId;
        }

        throw new InvalidOperationException("Aucun site SharePoint sélectionné.");
    }

    private async Task<GraphServiceClient> GetClientAsync()
        => await _authService.GetAuthenticatedClientAsync();

    private async Task<string> GetSiteIdForRequestAsync()
        => await GetSiteIdAsync();

    // Toutes les méthodes utilisent maintenant GetSiteIdForRequestAsync() au lieu de _siteId
}