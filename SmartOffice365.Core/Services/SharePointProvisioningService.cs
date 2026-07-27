using Microsoft.Graph;
using Microsoft.Graph.Models;
using SmartOffice365.Core.Interfaces;

namespace SmartOffice365.Core.Services
{
    public class SharePointProvisioningService : ISharePointProvisioningService
    {
        private readonly IGraphAuthService _authService;
        private readonly string _siteId;

        private static readonly Dictionary<string, List<ColumnDefinition>> ListDefinitions = new()
        {
            ["Contacts_Et_Entreprises"] = new List<ColumnDefinition>
            {
                new() { Name = "Role", Text = new TextColumn() },
                new() { Name = "Email", Text = new TextColumn() },
                new() { Name = "Telephone", Text = new TextColumn() },
                new() { Name = "CompteTeams", Text = new TextColumn() },
                new() { Name = "CodeVendorLIFNR", Text = new TextColumn() },
                new() { Name = "Entreprise", Text = new TextColumn() },
            },
            ["Affaires_et_Projets"] = new List<ColumnDefinition>
            {
                new() { Name = "CodeUniteSAP", Text = new TextColumn() },
                new() { Name = "DateDebutPrevue", DateTime = new DateTimeColumn() },
                new() { Name = "DateFinPrevue", DateTime = new DateTimeColumn() },
                new() { Name = "Statut", Choice = new ChoiceColumn { Choices = new List<string> { "Planifié", "En cours", "Terminé", "Annulé" } } },
                new() { Name = "Responsable", Text = new TextColumn() },
                new() { Name = "AvancementGlobal", Number = new NumberColumn { Minimum = 0, Maximum = 100 } },
            },
            ["Ordres_De_Travail"] = new List<ColumnDefinition>
            {
                new() { Name = "NumeroOT_Aufnr", Text = new TextColumn() },
                new() { Name = "NumeroEquipement_EQUNR", Text = new TextColumn() },
                new() { Name = "PosteTechnique_TPLNR", Text = new TextColumn() },
                new() { Name = "PosteTravail_ARBPL", Text = new TextColumn() },
                new() { Name = "Avancement", Number = new NumberColumn { Minimum = 0, Maximum = 100 } },
                new() { Name = "StatutShutdown", Choice = new ChoiceColumn { Choices = new List<string> { "À Faire", "En cours", "Bloqué", "Terminé" } } },
                new() { Name = "Priorite", Choice = new ChoiceColumn { Choices = new List<string> { "Critique", "Haute", "Normale", "Basse" } } },
                new() { Name = "Responsable", Text = new TextColumn() },
                new() { Name = "EntreprisePrestataire", Text = new TextColumn() },
                new() { Name = "DateDebutPrevue", DateTime = new DateTimeColumn() },
                new() { Name = "DateFinPrevue", DateTime = new DateTimeColumn() },
                new() { Name = "MotifsBlockage", Text = new TextColumn { AllowMultipleLines = true } },
            },
            ["Prerequis_et_Consignations"] = new List<ColumnDefinition>
            {
                new() { Name = "NumeroOT", Text = new TextColumn() },
                new() { Name = "Type", Choice = new ChoiceColumn { Choices = new List<string> { "Consignation électrique", "Permis de feu", "Permis de travail", "ATEX", "Travail en hauteur" } } },
                new() { Name = "EstValide", Boolean = new BooleanColumn() },
                new() { Name = "DateValidation", DateTime = new DateTimeColumn() },
                new() { Name = "Signataire", Text = new TextColumn() },
                new() { Name = "DateExpiration", DateTime = new DateTimeColumn() },
            },
            ["Ressources_Et_Moyens"] = new List<ColumnDefinition>
            {
                new() { Name = "NumeroOT", Text = new TextColumn() },
                new() { Name = "Type", Choice = new ChoiceColumn { Choices = new List<string> { "Main d'œuvre", "Matériel", "Outillage spécial", "Engin" } } },
                new() { Name = "EntreprisePrestataire", Text = new TextColumn() },
                new() { Name = "Description", Text = new TextColumn { AllowMultipleLines = true } },
                new() { Name = "QuantitePrevue", Number = new NumberColumn() },
                new() { Name = "QuantiteReelle", Number = new NumberColumn() },
                new() { Name = "Unite", Text = new TextColumn() },
                new() { Name = "EstDisponible", Boolean = new BooleanColumn() },
            },
            ["Habilitations_Contacts"] = new List<ColumnDefinition>
            {
                new() { Name = "NomContact", Text = new TextColumn() },
                new() { Name = "TypeHabilitation", Choice = new ChoiceColumn { Choices = new List<string> { "CACES R489", "CACES R482", "Habilitation Électrique B1", "Habilitation Électrique BR", "Travail en hauteur", "ATEX", "Pontier" } } },
                new() { Name = "Niveau", Text = new TextColumn() },
                new() { Name = "DateObtention", DateTime = new DateTimeColumn() },
                new() { Name = "DateExpiration", DateTime = new DateTimeColumn() },
                new() { Name = "Organisme", Text = new TextColumn() },
            }
        };

        public SharePointProvisioningService(IGraphAuthService authService, string siteId)
        {
            _authService = authService;
            _siteId = siteId;
        }

        public async Task<bool> TestSiteConnectionAsync()
        {
            try
            {
                var client = await _authService.GetAuthenticatedClientAsync();
                var site = await client.Sites[_siteId].GetAsync();
                return site != null;
            }
            catch { return false; }
        }

        public async Task<bool> ListExistsAsync(string listName)
        {
            try
            {
                var client = await _authService.GetAuthenticatedClientAsync();
                var lists = await client.Sites[_siteId].Lists.GetAsync(config =>
                    config.QueryParameters.Filter = $"displayName eq '{listName}'");
                return lists?.Value?.Any() == true;
            }
            catch { return false; }
        }

        public async Task<ProvisioningResult> ProvisionAllListsAsync(IProgress<string>? progress = null)
        {
            var result = new ProvisioningResult();
            try
            {
                var client = await _authService.GetAuthenticatedClientAsync();

                foreach (var (listName, columns) in ListDefinitions)
                {
                    progress?.Report($"Vérification de la liste '{listName}'...");

                    if (await ListExistsAsync(listName))
                    {
                        result.ActionsPerformed.Add($"✓ Liste '{listName}' existe déjà — ignorée.");
                        progress?.Report($"✓ '{listName}' existe déjà.");
                        continue;
                    }

                    progress?.Report($"⚙ Création de la liste '{listName}'...");
                    var newList = new Microsoft.Graph.Models.List
                    {
                        DisplayName = listName,
                        ListProp = new ListInfo { Template = "genericList" }
                    };

                    var createdList = await client.Sites[_siteId].Lists.PostAsync(newList);
                    result.ActionsPerformed.Add($"+ Liste '{listName}' créée (ID: {createdList?.Id})");

                    if (createdList?.Id == null) continue;

                    foreach (var column in columns)
                    {
                        await client.Sites[_siteId].Lists[createdList.Id].Columns.PostAsync(column);
                        result.ActionsPerformed.Add($"  + Colonne '{column.Name}' ajoutée à '{listName}'");
                    }
                    progress?.Report($"✓ Liste '{listName}' provisionnée avec {columns.Count} colonnes.");
                }

                result.Success = true;
                result.ActionsPerformed.Add("\n✅ Provisioning terminé avec succès.");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Errors.Add($"Erreur critique : {ex.Message}");
            }
            return result;
        }
    }
}
