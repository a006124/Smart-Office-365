using System;
using System.Threading.Tasks;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Interfaces
{
    /// <summary>
    /// Service de création automatique (provisioning) des listes SharePoint
    /// </summary>
    public interface ISharePointProvisioningService
    {
        /// <summary>
        /// Initialise le service avec l'ID ou l'URL du site sélectionné
        /// </summary>
        void SetSiteId(string siteId); // ◄--- AJOUTEZ CETTE LIGNE

        /// <summary>
        /// Teste la connexion au site SharePoint
        /// </summary>
        Task<bool> TestSiteConnectionAsync();

        /// <summary>
        /// Vérifie si une liste existe déjà sur le site
        /// </summary>
        Task<bool> ListExistsAsync(string listName);

        /// <summary>
        /// Lance la création de toutes les listes requises
        /// </summary>
        Task<SharePointProvisioningResult> ProvisionAllListsAsync(IProgress<string>? progress = null);
    }
}
