using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using SmartOffice365.Core.Interfaces;

namespace SmartOffice365.UI.ViewModels
{
    public class SharePointService : ISharePointService
    {
        private readonly IGraphAuthService _authService;
        private readonly HttpClient _httpClient;

        // TODO: Remplacez par l'ID de votre site SharePoint Teams (ex: "tenant.sharepoint.com,xxxx-xxxx,yyyy-yyyy")
        private const string SiteId = "VOTRE_SITE_SHAREPOINT_ID";
        private const string ListName = "Arrêts";

        public SharePointService(IGraphAuthService authService)
        {
            _authService = authService;
            _httpClient = new HttpClient();
        }

        private async Task PrepareHttpClientAsync()
        {
            // Récupère le token d'authentification de la session active
            string token = await _authService.GetAccessTokenAsync(); // Assurez-vous que cette méthode existe dans votre IGraphAuthService
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        public async Task<List<ArretModel>> GetArretsAsync()
        {
            await PrepareHttpClientAsync();

            // Endpoint Graph pour récupérer les éléments d'une liste SharePoint
            string url = $"https://graph.microsoft.com/v1.0/sites/{SiteId}/lists/{ListName}/items?expand=fields";
            var response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var valueArray = root.GetProperty("value");

            var list = new List<ArretModel>();
            foreach (var item in valueArray.EnumerateArray())
            {
                var fields = item.GetProperty("fields");

                list.Add(new ArretModel
                {
                    // L'ID SharePoint est stocké dans l'élément parent
                    Id = int.Parse(item.GetProperty("id").GetString()),
                    Titre = fields.TryGetProperty("Title", out var t) ? t.GetString() : "",
                    DateDebut = fields.TryGetProperty("DateDebut", out var dd) ? dd.GetDateTime() : DateTime.Now,
                    DateFin = fields.TryGetProperty("DateFin", out var df) ? df.GetDateTime() : DateTime.Now,
                    Statut = fields.TryGetProperty("Statut", out var s) ? s.GetString() : "En préparation",
                    Description = fields.TryGetProperty("Description", out var desc) ? desc.GetString() : "",
                    JalonsPreparation = fields.TryGetProperty("JalonsPreparation", out var j) ? j.GetString() : ""
                });
            }

            return list;
        }

        public async Task<ArretModel> CreateArretAsync(ArretModel arret)
        {
            await PrepareHttpClientAsync();

            string url = $"https://graph.microsoft.com/v1.0/sites/{SiteId}/lists/{ListName}/items";

            var payload = new
            {
                fields = new
                {
                    Title = arret.Titre,
                    DateDebut = arret.DateDebut.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    DateFin = arret.DateFin.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                    Statut = arret.Statut,
                    Description = arret.Description,
                    JalonsPreparation = arret.JalonsPreparation
                }
            };

            var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync(url, content);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            arret.Id = int.Parse(doc.RootElement.GetProperty("id").GetString());

            return arret;
        }

        public async Task UpdateArretAsync(ArretModel arret)
        {
            await PrepareHttpClientAsync();

            // Endpoint pour mettre à jour les champs d'un élément spécifique
            string url = $"https://graph.microsoft.com/v1.0/sites/{SiteId}/lists/{ListName}/items/{arret.Id}/fields";

            var payload = new
            {
                Title = arret.Titre,
                DateDebut = arret.DateDebut.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                DateFin = arret.DateFin.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                Statut = arret.Statut,
                Description = arret.Description,
                JalonsPreparation = arret.JalonsPreparation
            };

            var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

            // Utilisation de PATCH pour mettre à jour uniquement les champs spécifiés
            var request = new HttpRequestMessage(new HttpMethod("PATCH"), url) { Content = content };
            var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();
        }

        public async Task DeleteArretAsync(int id)
        {
            await PrepareHttpClientAsync();

            string url = $"https://graph.microsoft.com/v1.0/sites/{SiteId}/lists/{ListName}/items/{id}";
            var response = await _httpClient.DeleteAsync(url);
            response.EnsureSuccessStatusCode();
        }
    }
}
