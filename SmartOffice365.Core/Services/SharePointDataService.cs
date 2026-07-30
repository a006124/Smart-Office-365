using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;

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

    // --- IMPLÉMENTATION DE L'INTERFACE ISharePointDataService ---

    // Contacts
    public Task<List<ContactEntity>> GetContactsAsync(string? filter = null)
        => Task.FromResult(new List<ContactEntity>());

    public Task<ContactEntity?> GetContactByIdAsync(int id)
        => Task.FromResult<ContactEntity?>(null);

    public Task CreateContactAsync(ContactEntity contact)
        => Task.CompletedTask;

    public Task UpdateContactAsync(ContactEntity contact)
        => Task.CompletedTask;

    public Task DeleteContactAsync(int id)
        => Task.CompletedTask;

    // Affaires
    public Task<List<AffaireEntity>> GetAffairesAsync(string? filter = null)
        => Task.FromResult(new List<AffaireEntity>());

    public Task<AffaireEntity?> GetAffaireByIdAsync(int id)
        => Task.FromResult<AffaireEntity?>(null);

    public Task CreateAffaireAsync(AffaireEntity affaire)
        => Task.CompletedTask;

    public Task UpdateAffaireAsync(AffaireEntity affaire)
        => Task.CompletedTask;

    // Ordres de Travail
    public Task<List<OrdreDeTravailEntity>> GetOrdresDeTravailAsync(string? filter = null)
        => Task.FromResult(new List<OrdreDeTravailEntity>());

    public Task<OrdreDeTravailEntity?> GetOrdreDeTravailByIdAsync(int id)
        => Task.FromResult<OrdreDeTravailEntity?>(null);

    public Task CreateOrdreDeTravailAsync(OrdreDeTravailEntity ot)
        => Task.CompletedTask;

    public Task UpdateOrdreDeTravailAsync(OrdreDeTravailEntity ot)
        => Task.CompletedTask;

    public Task UpdateAvancementAsync(int id, int avancement, string statut)
        => Task.CompletedTask;

    // Prérequis
    public Task<List<PrerequsEntity>> GetPrerequisByOTAsync(int otId)
        => Task.FromResult(new List<PrerequsEntity>());

    public Task CreatePrerequsAsync(PrerequsEntity prereq)
        => Task.CompletedTask;

    public Task UpdatePrerequsAsync(PrerequsEntity prereq)
        => Task.CompletedTask;

    // Ressources
    public Task<List<RessourceEntity>> GetRessourcesByOTAsync(int otId)
        => Task.FromResult(new List<RessourceEntity>());

    public Task CreateRessourceAsync(RessourceEntity ressource)
        => Task.CompletedTask;

    public Task UpdateRessourceAsync(RessourceEntity ressource)
        => Task.CompletedTask;

    // Habilitations
    public Task<List<HabilitationEntity>> GetHabilitationsByContactAsync(int contactId)
        => Task.FromResult(new List<HabilitationEntity>());

    public Task CreateHabilitationAsync(HabilitationEntity habilitation)
        => Task.CompletedTask;

    // KPIs Dashboard
    public Task<DashboardKpis?> GetDashboardKpisAsync()
        => Task.FromResult<DashboardKpis?>(null);
}
