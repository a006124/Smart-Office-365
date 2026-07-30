using System.Collections.Generic;
using System.Threading.Tasks;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Interfaces
{
    /// <summary>
    /// Service CRUD pour les 6 listes SharePoint
    /// </summary>
    public interface ISharePointDataService
    {
        // Contacts
        Task<List<ContactEntity>> GetContactsAsync(string? filter = null);
        Task<ContactEntity?> GetContactByIdAsync(int id);
        Task CreateContactAsync(ContactEntity contact);
        Task UpdateContactAsync(ContactEntity contact);
        Task DeleteContactAsync(int id);

        // Affaires
        Task<List<AffaireEntity>> GetAffairesAsync(string? filter = null);
        Task<AffaireEntity?> GetAffaireByIdAsync(int id);
        Task CreateAffaireAsync(AffaireEntity affaire);
        Task UpdateAffaireAsync(AffaireEntity affaire);

        // Ordres de Travail
        Task<List<OrdreDeTravailEntity>> GetOrdresDeTravailAsync(string? filter = null);
        Task<OrdreDeTravailEntity?> GetOrdreDeTravailByIdAsync(int id);
        Task CreateOrdreDeTravailAsync(OrdreDeTravailEntity ot);
        Task UpdateOrdreDeTravailAsync(OrdreDeTravailEntity ot);
        Task UpdateAvancementAsync(int id, int avancement, string statut);

        // Prérequis
        Task<List<PrerequsEntity>> GetPrerequisByOTAsync(int otId);
        Task CreatePrerequsAsync(PrerequsEntity prereq);
        Task UpdatePrerequsAsync(PrerequsEntity prereq);

        // Ressources
        Task<List<RessourceEntity>> GetRessourcesByOTAsync(int otId);
        Task CreateRessourceAsync(RessourceEntity ressource);
        Task UpdateRessourceAsync(RessourceEntity ressource);

        // Habilitations
        Task<List<HabilitationEntity>> GetHabilitationsByContactAsync(int contactId);
        Task CreateHabilitationAsync(HabilitationEntity habilitation);

        // KPIs Dashboard
        Task<DashboardKpis?> GetDashboardKpisAsync(); // Ajustez le type de retour selon votre modèle de KPI
    }
}
