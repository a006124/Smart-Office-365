using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Interfaces
{
    /// <summary>
    /// Service de sélection du site SharePoint
    /// </summary>
    public interface ISharePointSelectionService
    {
        /// <summary>Obtient la liste des sites SharePoint disponibles</summary>
        Task<List<SharePointSiteInfo>> GetAvailableSitesAsync();

        /// <summary>Définit le site actif par son ID</summary>
        Task SetActiveSiteAsync(string siteId);

        /// <summary>Retourne l'ID du site actif</summary>
        string GetActiveSiteId();

        /// <summary>Vérifie si un site est sélectionné</summary>
        bool HasActiveSite();
    }
}