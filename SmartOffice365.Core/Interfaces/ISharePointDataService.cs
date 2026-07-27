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
        Task<ContactEntity> CreateContactAsync(ContactEntity contact);
        Task<ContactEntity> UpdateContactAsync(ContactEntity contact);
        Task DeleteContactAsync(int id);

        // Affaires
        Task<List<AffaireEntity>> GetAffairesAsync(string? filter = null);
        Task<AffaireEntity?> GetAffaireByIdAsync(int id);
        Task<AffaireEntity> CreateAffaireAsync(AffaireEntity affaire);
        Task<AffaireEntity> UpdateAffaireAsync(AffaireEntity affaire);

        // Ordres de Travail
        Task<List<OrdreDeTravailEntity>> GetOrdresDeTravailAsync(string? filter = null);
        Task<OrdreDeTravailEntity?> GetOrdreDeTravailByIdAsync(int id);
        Task<OrdreDeTravailEntity> CreateOrdreDeTravailAsync(OrdreDeTravailEntity ot);
        Task<OrdreDeTravailEntity> UpdateOrdreDeTravailAsync(OrdreDeTravailEntity ot);
        Task UpdateAvancementAsync(int id, int avancement, string statut);

        // Prérequis
        Task<List<PrerequsEntity>> GetPrerequisByOTAsync(int otId);
        Task<PrerequsEntity> CreatePrerequsAsync(PrerequsEntity prereq);
        Task<PrerequsEntity> UpdatePrerequsAsync(PrerequsEntity prereq);

        // Ressources
        Task<List<RessourceEntity>> GetRessourcesByOTAsync(int otId);
        Task<RessourceEntity> CreateRessourceAsync(RessourceEntity ressource);
        Task<RessourceEntity> UpdateRessourceAsync(RessourceEntity ressource);

        // Habilitations
        Task<List<HabilitationEntity>> GetHabilitationsByContactAsync(int contactId);
        Task<HabilitationEntity> CreateHabilitationAsync(HabilitationEntity habilitation);

        // KPIs Dashboard
        Task<DashboardKpis> GetDashboardKpisAsync();
    }
}
