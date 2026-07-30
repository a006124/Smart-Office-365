// language: csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions; // ApiException (Graph v5)
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Services
{
    public class SharePointProvisioningService : ISharePointProvisioningService
    {
        private readonly IGraphAuthService _authService;
        private readonly ISharePointSelectionService? _selectionService;
        private string? _siteId; // Peut contenir une URL ou un ID Graph

        // Cache du jeton pour REST SharePoint
        private string? _cachedSharePointToken;

        // Tenant Renault (adapter si vous le rendez dynamique)
        private const string RenaultTenantId = "d6b0bbee-7cd9-4d60-bce6-4a67b543e2ae";

        public SharePointProvisioningService(
            IGraphAuthService authService,
            ISharePointSelectionService selectionService)
        {
            _authService = authService;
            _selectionService = selectionService;
        }

        public SharePointProvisioningService(IGraphAuthService authService, string siteId)
        {
            _authService = authService;
            _siteId = siteId;
            _selectionService = null;
        }

        public void SetSiteId(string siteId)
        {
            _siteId = siteId;
            _cachedSharePointToken = null; // Invalide le cache si la cible change
        }

        // ------------------ Résolution du Site ------------------

        private async Task<string> GetSiteIdAsync()
        {
            if (string.IsNullOrEmpty(_siteId) && _selectionService?.HasActiveSite() == true)
                _siteId = _selectionService.GetActiveSiteId();

            if (string.IsNullOrEmpty(_siteId))
                throw new InvalidOperationException("Aucun site SharePoint sélectionné. Saisissez une URL ou choisissez un site actif.");

            // Si _siteId est une URL, on la résout en ID via Graph
            if (_siteId.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                _siteId = await ResolveUrlToIdAsync(_siteId);

            return _siteId!;
        }

        private async Task<string> ResolveUrlToIdAsync(string url)
        {
            var client = await _authService.GetAuthenticatedClientAsync();
            var cleanUrl = url.Trim().TrimEnd('/');
            var uri = new Uri(cleanUrl);
            var host = uri.Host;         // ex: grouperenault.sharepoint.com
            var path = uri.AbsolutePath; // ex: /sites/MonSite

            var site = await client.Sites[$"{host}:{path}"].GetAsync();
            if (!string.IsNullOrEmpty(site?.Id))
                return site!.Id!;
            throw new Exception("Le site SharePoint est introuvable (Graph). Vérifiez l'URL.");
        }

        // Convertit un ID Graph en WebUrl, ou renvoie l’URL telle quelle si déjà URL
        private async Task<string> EnsureWebUrlAsync(string siteIdOrUrl)
        {
            if (siteIdOrUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                return siteIdOrUrl.Trim().TrimEnd('/');

            var client = await _authService.GetAuthenticatedClientAsync();
            var site = await client.Sites[siteIdOrUrl].GetAsync();
            if (!string.IsNullOrEmpty(site?.WebUrl))
                return site!.WebUrl!.Trim().TrimEnd('/');
            throw new Exception("Impossible de récupérer l’URL du site depuis Graph.");
        }

        // ------------------ Définition Listes / Colonnes ------------------

        private static readonly Dictionary<string, List<ColumnDef>> ListDefinitions = new()
        {
            ["Contacts_Et_Entreprises"] = new List<ColumnDef>
            {
                new("Role", "Text"),
                new("Email", "Text"),
                new("Telephone", "Text"),
                new("CompteTeams", "Text"),
                new("CodeVendorLIFNR", "Text"),
                new("Entreprise", "Text")
            },
            ["Affaires_et_Projets"] = new List<ColumnDef>
            {
                new("CodeUniteSAP", "Text"),
                new("DateDebutPrevue", "DateTime"),
                new("DateFinPrevue", "DateTime"),
                new("Statut", "Choice", new[] { "Planifié", "En cours", "Terminé", "Annulé" }),
                new("Responsable", "Text"),
                new("AvancementGlobal", "Number")
            },
            ["Ordres_De_Travail"] = new List<ColumnDef>
            {
                new("NumeroOT_Aufnr", "Text"),
                new("NumeroEquipement_EQUNR", "Text"),
                new("PosteTechnique_TPLNR", "Text"),
                new("PosteTravail_ARBPL", "Text"),
                new("Avancement", "Number"),
                new("StatutShutdown", "Choice", new[] { "À Faire", "En cours", "Bloqué", "Terminé" }),
                new("Priorite", "Choice", new[] { "Critique", "Haute", "Normale", "Basse" }),
                new("Responsable", "Text"),
                new("EntreprisePrestataire", "Text"),
                new("DateDebutPrevue", "DateTime"),
                new("DateFinPrevue", "DateTime"),
                new("MotifsBlockage", "Note")
            },
            ["Prerequis_et_Consignations"] = new List<ColumnDef>
            {
                new("NumeroOT", "Text"),
                new("Type", "Choice", new[] { "Consignation électrique", "Permis de feu", "Permis de travail", "ATEX", "Travail en hauteur" }),
                new("EstValide", "Boolean"),
                new("DateValidation", "DateTime"),
                new("Signataire", "Text"),
                new("DateExpiration", "DateTime")
            },
            ["Ressources_Et_Moyens"] = new List<ColumnDef>
            {
                new("NumeroOT", "Text"),
                new("Type", "Choice", new[] { "Main d'œuvre", "Matériel", "Outillage spécial", "Engin" }),
                new("EntreprisePrestataire", "Text"),
                new("Description", "Note"),
                new("QuantitePrevue", "Number"),
                new("QuantiteReelle", "Number"),
                new("Unite", "Text"),
                new("EstDisponible", "Boolean")
            },
            ["Habilitations_Contacts"] = new List<ColumnDef>
            {
                new("NomContact", "Text"),
                new("TypeHabilitation", "Choice", new[] { "CACES R489", "CACES R482", "Habilitation Électrique B1", "Habilitation Électrique BR", "Travail en hauteur", "ATEX", "Pontier" }),
                new("Niveau", "Text"),
                new("DateObtention", "DateTime"),
                new("DateExpiration", "DateTime"),
                new("Organisme", "Text")
            }
        };

        // ------------------ Tests ------------------

        public async Task<bool> TestSiteConnectionAsync()
        {
            // Essai via Graph
            try
            {
                var siteId = await GetSiteIdAsync();
                var client = await _authService.GetAuthenticatedClientAsync();
                var site = await client.Sites[siteId].GetAsync();
                return site != null;
            }
            catch
            {
                // Fallback REST (GET _api/web)
                try
                {
                    var siteUrl = await EnsureWebUrlAsync(await GetSiteIdAsync());
                    var token = await GetSharePointAccessTokenAsync(siteUrl);
                    using var http = new HttpClient();
                    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var resp = await http.GetAsync($"{siteUrl}/_api/web");
                    return resp.IsSuccessStatusCode;
                }
                catch { return false; }
            }
        }

        public async Task<bool> ListExistsAsync(string listName)
        {
            // Essai via Graph
            try
            {
                var siteId = await GetSiteIdAsync();
                var client = await _authService.GetAuthenticatedClientAsync();
                var lists = await client.Sites[siteId].Lists.GetAsync(cfg =>
                    cfg.QueryParameters.Filter = $"displayName eq '{listName}'");
                return lists?.Value?.Any() == true;
            }
            catch (ApiException ex) when (ex.ResponseStatusCode == 401 || ex.ResponseStatusCode == 403)
            {
                // Fallback REST si Graph refuse l’accès
            }
            catch
            {
                // Autres erreurs Graph : on tentera aussi REST
            }

            // Fallback REST
            try
            {
                var siteUrl = await EnsureWebUrlAsync(await GetSiteIdAsync());
                var token = await GetSharePointAccessTokenAsync(siteUrl);
                using var http = new HttpClient();
                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                http.DefaultRequestHeaders.Accept.Clear();
                http.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=verbose");

                var resp = await http.GetAsync($"{siteUrl}/_api/web/lists/getByTitle('{listName}')");
                return resp.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        // ------------------ Provisionning (Graph + Fallback REST) ------------------

        public async Task<SharePointProvisioningResult> ProvisionAllListsAsync(IProgress<string>? progress = null)
        {
            var result = new SharePointProvisioningResult();

            try
            {
                var siteId = await GetSiteIdAsync();
                var client = await _authService.GetAuthenticatedClientAsync();
                progress?.Report("Connexion sécurisée à Microsoft Graph...");

                foreach (var (listName, columns) in ListDefinitions)
                {
                    progress?.Report($"Vérification de la liste '{listName}'...");

                    if (await ListExistsAsync(listName))
                    {
                        result.ActionsPerformed.Add($"✓ Liste '{listName}' existe déjà — ignorée.");
                        progress?.Report($"✓ '{listName}' existe déjà.");
                        continue;
                    }

                    // 1) Tentative via Graph
                    try
                    {
                        progress?.Report($"⚙ Création (Graph) de la liste '{listName}'...");
                        var newList = new Microsoft.Graph.Models.List
                        {
                            DisplayName = listName,
                            ListProp = new ListInfo { Template = "genericList" }
                        };

                        var createdList = await client.Sites[siteId].Lists.PostAsync(newList);
                        result.ActionsPerformed.Add($"+ Liste '{listName}' créée via Graph.");

                        if (!string.IsNullOrEmpty(createdList?.Id))
                        {
                            foreach (var col in columns)
                            {
                                var colDef = new ColumnDefinition
                                {
                                    Name = col.Name,
                                    DisplayName = col.Name
                                };

                                switch (col.Type.ToLower())
                                {
                                    case "datetime": colDef.DateTime = new DateTimeColumn(); break;
                                    case "number": colDef.Number = new NumberColumn(); break;
                                    case "boolean": colDef.Boolean = new BooleanColumn(); break;
                                    case "note": colDef.Text = new TextColumn { AllowMultipleLines = true }; break;
                                    case "choice": colDef.Choice = new ChoiceColumn { Choices = col.Choices?.ToList() }; break;
                                    default: colDef.Text = new TextColumn(); break;
                                }

                                await client.Sites[siteId].Lists[createdList.Id].Columns.PostAsync(colDef);
                                result.ActionsPerformed.Add($"  + Colonne '{col.Name}' (Graph) ajoutée.");
                            }
                        }

                        progress?.Report($"✓ Liste '{listName}' provisionnée (Graph).");
                        continue; // Liste suivante
                    }
                    catch (ApiException ex) when (ex.ResponseStatusCode == 401 || ex.ResponseStatusCode == 403)
                    {
                        // 2) Fallback REST si Graph renvoie 401/403
                        progress?.Report($"⚠ Graph Access Denied pour '{listName}'. Fallback via REST SharePoint...");
                        var siteUrl = await EnsureWebUrlAsync(siteId);
                        var ok = await CreateListAndColumnsViaRestAsync(siteUrl, listName, columns, progress, result);
                        if (!ok)
                            throw new Exception($"Échec du fallback REST sur '{listName}'.");
                        continue;
                    }
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

        // ------------------ Fallback REST Helpers ------------------

        private async Task<bool> CreateListAndColumnsViaRestAsync(
            string siteUrlParam,
            string listName,
            List<ColumnDef> columns,
            IProgress<string>? progress,
            SharePointProvisioningResult result,
            CancellationToken ct = default)
        {
            try
            {
                var siteUrl = siteUrlParam.TrimEnd('/');

                var token = await GetSharePointAccessTokenAsync(siteUrl, ct);

                // Sanity check d’accès
                using (var ping = new HttpClient())
                {
                    ping.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    ping.DefaultRequestHeaders.Accept.Clear();
                    ping.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=verbose");
                    var pingResp = await ping.GetAsync($"{siteUrl}/_api/web", ct);
                    var pingBody = await ReadBodyAsync(pingResp.Content, ct);
                    if (!pingResp.IsSuccessStatusCode)
                    {
                        LogHttpFailure(progress, pingResp, pingBody, "GET /_api/web");
                        throw new Exception($"Accès Web refusé: {ExtractSharePointErrorMessage(pingBody)}");
                    }
                }

                // Récupère le FormDigest (obligatoire pour POST REST)
                var formDigest = await GetFormDigestAsync(siteUrl, token, progress, ct);

                using var http = new HttpClient();
                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                http.DefaultRequestHeaders.Accept.Clear();
                http.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=verbose");
                http.DefaultRequestHeaders.TryAddWithoutValidation("X-RequestDigest", formDigest);

                // Création de la liste
                var listPayload = new
                {
                    __metadata = new { type = "SP.List" },
                    AllowContentTypes = true,
                    BaseTemplate = 100,
                    Title = listName
                };
                using var listContent = new StringContent(JsonSerializer.Serialize(listPayload), Encoding.UTF8);
                listContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");

                var respList = await http.PostAsync($"{siteUrl}/_api/web/lists", listContent, ct);
                var respListBody = await ReadBodyAsync(respList.Content, ct);
                if (!respList.IsSuccessStatusCode)
                {
                    LogHttpFailure(progress, respList, respListBody, $"POST create list {listName}");
                    throw new Exception(ExtractSharePointErrorMessage(respListBody));
                }
                result.ActionsPerformed.Add($"+ Liste '{listName}' créée (REST).");

                // Ajout des colonnes
                foreach (var col in columns)
                {
                    var schemaXml = BuildColumnXmlSchema(col);
                    var columnPayload = new
                    {
                        parameters = new
                        {
                            __metadata = new { type = "SP.XmlSchemaFieldCreationInformation" },
                            SchemaXml = schemaXml
                        }
                    };
                    using var colContent = new StringContent(JsonSerializer.Serialize(columnPayload), Encoding.UTF8);
                    colContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");

                    var respCol = await http.PostAsync($"{siteUrl}/_api/web/lists/getByTitle('{listName}')/fields/createfieldasxml", colContent, ct);
                    var respColBody = await ReadBodyAsync(respCol.Content, ct);

                    if (!respCol.IsSuccessStatusCode)
                    {
                        LogHttpFailure(progress, respCol, respColBody, $"POST add column {col.Name}");
                        progress?.Report($"  ⚠ Erreur colonne '{col.Name}' : {ExtractSharePointErrorMessage(respColBody)}");
                    }
                    else
                    {
                        result.ActionsPerformed.Add($"  + Colonne '{col.Name}' (REST) ajoutée.");
                    }
                }

                progress?.Report($"✓ Liste '{listName}' provisionnée (REST).");
                return true;
            }
            catch (Exception ex)
            {
                progress?.Report($"❌ Fallback REST : {ex.Message}");
                return false;
            }
        }

        private async Task<string> GetFormDigestAsync(string siteUrl, string bearerToken, IProgress<string>? progress, CancellationToken ct = default)
        {
            using var http = new HttpClient();
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
            http.DefaultRequestHeaders.Accept.Clear();
            http.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=verbose");
            http.DefaultRequestHeaders.TryAddWithoutValidation("X-RequestForceAuthentication", "true");

            // POST sans corps (Content-Length: 0)
            using var req = new HttpRequestMessage(HttpMethod.Post, $"{siteUrl}/_api/contextinfo");
            req.Headers.TryAddWithoutValidation("Content-Type", "application/json;odata=verbose");

            var resp = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
            var payload = await ReadBodyAsync(resp.Content, ct);

            if (!resp.IsSuccessStatusCode)
            {
                LogHttpFailure(progress, resp, payload, "POST /_api/contextinfo");
                // Fallback "nometadata" si le tenant n’accepte pas verbose
                http.DefaultRequestHeaders.Remove("Accept");
                http.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");
                using var req2 = new HttpRequestMessage(HttpMethod.Post, $"{siteUrl}/_api/contextinfo");
                req2.Headers.TryAddWithoutValidation("Content-Type", "application/json;odata=nometadata");
                var resp2 = await http.SendAsync(req2, HttpCompletionOption.ResponseHeadersRead, ct);
                var payload2 = await ReadBodyAsync(resp2.Content, ct);
                if (!resp2.IsSuccessStatusCode)
                {
                    LogHttpFailure(progress, resp2, payload2, "POST /_api/contextinfo (nometadata)");
                    throw new Exception($"Échec contextinfo : {ExtractSharePointErrorMessage(payload2)}");
                }
                payload = payload2;
            }

            using var doc = JsonDocument.Parse(payload);
            string? digest = null;

            // verbose
            if (doc.RootElement.TryGetProperty("d", out var dObj) &&
                dObj.TryGetProperty("GetContextWebInformation", out var gci) &&
                gci.TryGetProperty("FormDigestValue", out var fdv))
            {
                digest = fdv.GetString();
            }
            // nometadata (parfois renvoie directement FormDigestValue)
            if (string.IsNullOrWhiteSpace(digest) &&
                doc.RootElement.TryGetProperty("FormDigestValue", out var fdv2))
            {
                digest = fdv2.GetString();
            }

            if (string.IsNullOrWhiteSpace(digest))
                throw new Exception("FormDigestValue vide ou introuvable.");
            return digest!;
        }

        private async Task<string> GetSharePointAccessTokenAsync(string siteUrl, CancellationToken ct = default)
        {
            if (!string.IsNullOrEmpty(_cachedSharePointToken))
                return _cachedSharePointToken!;

            var uri = new Uri(siteUrl);
            var resource = $"{uri.Scheme}://{uri.Host}"; // ex: https://grouperenault.sharepoint.com

            // Acquisition robuste : SharedTokenCache → VisualStudio → InteractiveBrowser
            var chained = new ChainedTokenCredential(
                new SharedTokenCacheCredential(new SharedTokenCacheCredentialOptions
                {
                    TenantId = RenaultTenantId
                    // Username = "prenom.nom@renault.com" // Optionnel si vous souhaitez cibler 1 profil précis
                }),
                new VisualStudioCredential(new VisualStudioCredentialOptions
                {
                    TenantId = RenaultTenantId
                }),
                new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
                {
                    TenantId = RenaultTenantId
                })
            );

            var ctx = new TokenRequestContext(new[] { $"{resource}/.default" });
            var token = await chained.GetTokenAsync(ctx, ct);
            _cachedSharePointToken = token.Token;
            return _cachedSharePointToken!;
        }

        // ------------------ Utilitaires Log/Erreurs/XML ------------------

        private async Task<string> ReadBodyAsync(HttpContent content, CancellationToken ct)
        {
            try { return await content.ReadAsStringAsync(ct); }
            catch { return string.Empty; }
        }

        private void LogHttpFailure(IProgress<string>? progress, HttpResponseMessage resp, string body, string context)
        {
            resp.Headers.TryGetValues("SPRequestGuid", out var spGuid);
            resp.Headers.TryGetValues("request-id", out var reqId);
            var guid = spGuid?.FirstOrDefault() ?? reqId?.FirstOrDefault() ?? "(n/a)";
            progress?.Report($"❌ {context} — HTTP {(int)resp.StatusCode} {resp.ReasonPhrase} — SPRequestGuid: {guid}");
            if (!string.IsNullOrWhiteSpace(body))
                progress?.Report($"Détails SharePoint: {ExtractSharePointErrorMessage(body)}");
        }

        private string ExtractSharePointErrorMessage(string jsonResponse)
        {
            try
            {
                using var doc = JsonDocument.Parse(jsonResponse);
                if (doc.RootElement.TryGetProperty("error", out var e) &&
                    e.TryGetProperty("message", out var m) &&
                    (m.TryGetProperty("value", out var v) || m.ValueKind == JsonValueKind.String))
                {
                    if (v.ValueKind != JsonValueKind.Undefined) return v.GetString() ?? jsonResponse;
                    return m.GetString() ?? jsonResponse;
                }
            }
            catch { }
            return jsonResponse;
        }

        private string BuildColumnXmlSchema(ColumnDef col)
        {
            var name = col.Name;
            return col.Type.ToLower() switch
            {
                "datetime" => $"<Field Type='DateTime' DisplayName='{name}' Name='{name}' Format='DateOnly' />",
                "number" => $"<Field Type='Number' DisplayName='{name}' Name='{name}' Decimals='0' />",
                "boolean" => $"<Field Type='Boolean' DisplayName='{name}' Name='{name}' />",
                "note" => $"<Field Type='Note' DisplayName='{name}' Name='{name}' NumLines='6' RichText='FALSE' />",
                "choice" => BuildChoiceXml(name, col.Choices),
                _ => $"<Field Type='Text' DisplayName='{name}' Name='{name}' />"
            };
        }

        private string BuildChoiceXml(string name, string[]? choices)
        {
            var sb = new StringBuilder();
            sb.Append($"<Field Type='Choice' DisplayName='{name}' Name='{name}'><CHOICES>");
            foreach (var c in choices ?? Array.Empty<string>())
                sb.Append($"<CHOICE>{c}</CHOICE>");
            sb.Append("</CHOICES></Field>");
            return sb.ToString();
        }

        // Modèle interne
        private class ColumnDef
        {
            public string Name { get; }
            public string Type { get; }
            public string[]? Choices { get; }

            public ColumnDef(string name, string type, string[]? choices = null)
            {
                Name = name;
                Type = type;
                Choices = choices;
            }
        }
    }
}
