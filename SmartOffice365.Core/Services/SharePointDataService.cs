using Microsoft.Graph;
using Microsoft.Graph.Models;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Services
{
    public class SharePointDataService : ISharePointDataService
    {
        private readonly IGraphAuthService _authService;
        private readonly string _siteId;

        public SharePointDataService(IGraphAuthService authService, string siteId)
        {
            _authService = authService;
            _siteId = siteId;
        }

        private async Task<GraphServiceClient> GetClientAsync()
            => await _authService.GetAuthenticatedClientAsync();

        // ====== CONTACTS ======
        public async Task<List<ContactEntity>> GetContactsAsync(string? filter = null)
        {
            var client = await GetClientAsync();
            var items = await client.Sites[_siteId].Lists["Contacts_Et_Entreprises"].Items
                .GetAsync(config => { config.QueryParameters.Expand = new[] { "fields" }; });
            return items?.Value?.Select(MapToContact).ToList() ?? new List<ContactEntity>();
        }

        public async Task<ContactEntity?> GetContactByIdAsync(int id)
        {
            var contacts = await GetContactsAsync();
            return contacts.FirstOrDefault(c => c.Id == id);
        }

        public async Task<ContactEntity> CreateContactAsync(ContactEntity contact)
        {
            var client = await GetClientAsync();
            var fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["Title"] = $"{contact.Prenom} {contact.Nom}",
                    ["Role"] = contact.Role,
                    ["Email"] = contact.Email,
                    ["Telephone"] = contact.Telephone,
                    ["CompteTeams"] = contact.CompteTeams,
                    ["CodeVendorLIFNR"] = contact.CodeVendorSapLIFNR,
                    ["Entreprise"] = contact.Entreprise,
                }
            };
            await client.Sites[_siteId].Lists["Contacts_Et_Entreprises"].Items.PostAsync(
                new ListItem { Fields = fields });
            return contact;
        }

        public async Task<ContactEntity> UpdateContactAsync(ContactEntity contact)
        {
            var client = await GetClientAsync();
            var fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["Role"] = contact.Role,
                    ["Email"] = contact.Email,
                    ["Telephone"] = contact.Telephone,
                }
            };
            await client.Sites[_siteId].Lists["Contacts_Et_Entreprises"].Items[contact.SharePointItemId].Fields.PatchAsync(fields);
            return contact;
        }

        public async Task DeleteContactAsync(int id)
        {
            var contact = await GetContactByIdAsync(id);
            if (contact == null) return;
            var client = await GetClientAsync();
            await client.Sites[_siteId].Lists["Contacts_Et_Entreprises"].Items[contact.SharePointItemId].DeleteAsync();
        }

        // ====== AFFAIRES ======
        public async Task<List<AffaireEntity>> GetAffairesAsync(string? filter = null)
        {
            var client = await GetClientAsync();
            var items = await client.Sites[_siteId].Lists["Affaires_et_Projets"].Items
                .GetAsync(config => { config.QueryParameters.Expand = new[] { "fields" }; });
            return items?.Value?.Select(MapToAffaire).ToList() ?? new List<AffaireEntity>();
        }

        public async Task<AffaireEntity?> GetAffaireByIdAsync(int id)
        {
            var affaires = await GetAffairesAsync();
            return affaires.FirstOrDefault(a => a.Id == id);
        }

        public async Task<AffaireEntity> CreateAffaireAsync(AffaireEntity affaire)
        {
            var client = await GetClientAsync();
            var fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["Title"] = affaire.Titre,
                    ["CodeUniteSAP"] = affaire.CodeUniteSAP,
                    ["Statut"] = affaire.Statut,
                    ["Responsable"] = affaire.Responsable,
                }
            };
            await client.Sites[_siteId].Lists["Affaires_et_Projets"].Items.PostAsync(
                new ListItem { Fields = fields });
            return affaire;
        }

        public async Task<AffaireEntity> UpdateAffaireAsync(AffaireEntity affaire)
            => affaire;

        // ====== ORDRES DE TRAVAIL ======
        public async Task<List<OrdreDeTravailEntity>> GetOrdresDeTravailAsync(string? filter = null)
        {
            var client = await GetClientAsync();
            var items = await client.Sites[_siteId].Lists["Ordres_De_Travail"].Items
                .GetAsync(config => { config.QueryParameters.Expand = new[] { "fields" }; });
            return items?.Value?.Select(MapToOT).ToList() ?? new List<OrdreDeTravailEntity>();
        }

        public async Task<OrdreDeTravailEntity?> GetOrdreDeTravailByIdAsync(int id)
        {
            var ots = await GetOrdresDeTravailAsync();
            return ots.FirstOrDefault(o => o.Id == id);
        }

        public async Task<OrdreDeTravailEntity> CreateOrdreDeTravailAsync(OrdreDeTravailEntity ot)
        {
            var client = await GetClientAsync();
            var fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["Title"] = ot.Titre,
                    ["NumeroOT_Aufnr"] = ot.NumeroOT_Aufnr,
                    ["NumeroEquipement_EQUNR"] = ot.NumeroEquipement_EQUNR,
                    ["PosteTechnique_TPLNR"] = ot.PosteTechnique_TPLNR,
                    ["PosteTravail_ARBPL"] = ot.PosteTravail_ARBPL,
                    ["Avancement"] = ot.Avancement,
                    ["StatutShutdown"] = ot.StatutShutdown,
                    ["Priorite"] = ot.Priorite,
                    ["Responsable"] = ot.Responsable,
                    ["EntreprisePrestataire"] = ot.EntreprisePrestataire,
                }
            };
            await client.Sites[_siteId].Lists["Ordres_De_Travail"].Items.PostAsync(
                new ListItem { Fields = fields });
            return ot;
        }

        public async Task<OrdreDeTravailEntity> UpdateOrdreDeTravailAsync(OrdreDeTravailEntity ot)
            => ot;

        public async Task UpdateAvancementAsync(int id, int avancement, string statut)
        {
            var ot = await GetOrdreDeTravailByIdAsync(id);
            if (ot == null) return;
            var client = await GetClientAsync();
            var fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["Avancement"] = avancement,
                    ["StatutShutdown"] = statut,
                }
            };
            await client.Sites[_siteId].Lists["Ordres_De_Travail"].Items[ot.SharePointItemId].Fields.PatchAsync(fields);
        }

        public async Task<List<PrerequsEntity>> GetPrerequisByOTAsync(int otId)
            => new List<PrerequsEntity>();

        public async Task<PrerequsEntity> CreatePrerequsAsync(PrerequsEntity prereq)
            => prereq;

        public async Task<PrerequsEntity> UpdatePrerequsAsync(PrerequsEntity prereq)
            => prereq;

        public async Task<List<RessourceEntity>> GetRessourcesByOTAsync(int otId)
            => new List<RessourceEntity>();

        public async Task<RessourceEntity> CreateRessourceAsync(RessourceEntity ressource)
            => ressource;

        public async Task<RessourceEntity> UpdateRessourceAsync(RessourceEntity ressource)
            => ressource;

        public async Task<List<HabilitationEntity>> GetHabilitationsByContactAsync(int contactId)
            => new List<HabilitationEntity>();

        public async Task<HabilitationEntity> CreateHabilitationAsync(HabilitationEntity habilitation)
            => habilitation;

        public async Task<DashboardKpis> GetDashboardKpisAsync()
        {
            var ots = await GetOrdresDeTravailAsync();
            return new DashboardKpis
            {
                TotalOT = ots.Count,
                OTTermines = ots.Count(o => o.StatutShutdown == "Terminé"),
                OTEnCours = ots.Count(o => o.StatutShutdown == "En cours"),
                OTBloques = ots.Count(o => o.StatutShutdown == "Bloqué"),
                OTEnRetard = ots.Count(o => o.StatutShutdown != "Terminé" && o.DateFinPrevue < DateTime.Now),
                AvancementGlobal = ots.Any() ? ots.Average(o => o.Avancement) : 0,
                TotalContacts = (await GetContactsAsync()).Count,
                TotalAffaires = (await GetAffairesAsync()).Count
            };
        }

        // ====== HELPERS & MAPPINGS ======
        private static string GetFieldString(IDictionary<string, object>? fields, string key, string defaultValue = "")
        {
            if (fields != null && fields.TryGetValue(key, out var val) && val != null)
            {
                return val.ToString() ?? defaultValue;
            }
            return defaultValue;
        }

        private static int GetFieldInt(IDictionary<string, object>? fields, string key, int defaultValue = 0)
        {
            if (fields != null && fields.TryGetValue(key, out var val) && val != null)
            {
                if (int.TryParse(val.ToString(), out int result)) return result;
            }
            return defaultValue;
        }

        private ContactEntity MapToContact(ListItem item)
        {
            var f = item.Fields?.AdditionalData;
            return new ContactEntity
            {
                SharePointItemId = item.Id ?? "",
                Nom = GetFieldString(f, "Title"),
                Role = GetFieldString(f, "Role"),
                Email = GetFieldString(f, "Email"),
                Telephone = GetFieldString(f, "Telephone"),
                CompteTeams = GetFieldString(f, "CompteTeams"),
                CodeVendorSapLIFNR = GetFieldString(f, "CodeVendorLIFNR"),
                Entreprise = GetFieldString(f, "Entreprise"),
            };
        }

        private AffaireEntity MapToAffaire(ListItem item)
        {
            var f = item.Fields?.AdditionalData;
            return new AffaireEntity
            {
                SharePointItemId = item.Id ?? "",
                Titre = GetFieldString(f, "Title"),
                CodeUniteSAP = GetFieldString(f, "CodeUniteSAP"),
                Statut = GetFieldString(f, "Statut"),
                Responsable = GetFieldString(f, "Responsable"),
            };
        }

        private OrdreDeTravailEntity MapToOT(ListItem item)
        {
            var f = item.Fields?.AdditionalData;
            return new OrdreDeTravailEntity
            {
                SharePointItemId = item.Id ?? "",
                Titre = GetFieldString(f, "Title"),
                NumeroOT_Aufnr = GetFieldString(f, "NumeroOT_Aufnr"),
                NumeroEquipement_EQUNR = GetFieldString(f, "NumeroEquipement_EQUNR"),
                PosteTechnique_TPLNR = GetFieldString(f, "PosteTechnique_TPLNR"),
                PosteTravail_ARBPL = GetFieldString(f, "PosteTravail_ARBPL"),
                Avancement = GetFieldInt(f, "Avancement", 0),
                StatutShutdown = GetFieldString(f, "StatutShutdown", "À Faire"),
                Priorite = GetFieldString(f, "Priorite", "Normale"),
                Responsable = GetFieldString(f, "Responsable"),
                EntreprisePrestataire = GetFieldString(f, "EntreprisePrestataire"),
            };
        }
    }
}
