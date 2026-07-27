using System;
using System.Threading.Tasks;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Interfaces
{
    /// <summary>
    /// Service de provisioning automatique des listes SharePoint
    /// </summary>
    public interface ISharePointProvisioningService
    {
        /// <summary>Vérifie l'accessibilité du site SharePoint configuré</summary>
        Task<bool> TestSiteConnectionAsync();

        /// <summary>Crée les 6 listes SharePoint requises si elles n'existent pas</summary>
        Task<SharePointProvisioningResult> ProvisionAllListsAsync(IProgress<string>? progress = null);

        /// <summary>Vérifie l'existence d'une liste par son nom</summary>
        Task<bool> ListExistsAsync(string listName);
    }
}
