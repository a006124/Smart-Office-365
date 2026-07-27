using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Interfaces
{
    /// <summary>
    /// Résultat d'une opération de provisioning SharePoint
    /// </summary>
    public class ProvisioningResult
    {
        public bool Success { get; set; }
        public List<string> ActionsPerformed { get; set; } = new();
        public List<string> Errors { get; set; } = new();
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    /// <summary>
    /// Service de provisioning automatique des listes SharePoint
    /// </summary>
    public interface ISharePointProvisioningService
    {
        /// <summary>Vérifie l'accessibilité du site SharePoint configuré</summary>
        Task<bool> TestSiteConnectionAsync();

        /// <summary>Crée les 6 listes SharePoint requises si elles n'existent pas</summary>
        Task<ProvisioningResult> ProvisionAllListsAsync(IProgress<string>? progress = null);

        /// <summary>Vérifie l'existence d'une liste par son nom</summary>
        Task<bool> ListExistsAsync(string listName);
    }
}
