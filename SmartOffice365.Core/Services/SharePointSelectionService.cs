// language: csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;

namespace SmartOffice365.Core.Services
{
    public class SharePointSelectionService : ISharePointSelectionService
    {
        private readonly IGraphAuthService _authService;

        // Valeur en mémoire
        private string? _activeSiteId;

        // Chemin du fichier de persistance (ex: C:\Users\<vous>\AppData\Roaming\SmartOffice365\activeSite.json)
        private readonly string _storePath;

        // Verrou pour rendre les accès thread-safe
        private readonly SemaphoreSlim _lock = new(1, 1);

        // Modèle de stockage (extensible si vous voulez mémoriser DisplayName/WebUrl par la suite)
        private class StoreModel
        {
            public string SiteId { get; set; } = string.Empty;
            public DateTime SavedAtUtc { get; set; } = DateTime.UtcNow;
        }

        public SharePointSelectionService(IGraphAuthService authService)
        {
            _authService = authService;

            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var folder = Path.Combine(appData, "SmartOffice365");
            _storePath = Path.Combine(folder, "activeSite.json");

            // Chargement à la construction (silent fail si le fichier n'existe pas)
            try
            {
                Directory.CreateDirectory(folder);
                LoadFromDisk();
            }
            catch
            {
                // En cas d'erreur d'I/O, on ignore et on repart sans site actif
                _activeSiteId = null;
            }
        }

        public async Task<List<SharePointSiteInfo>> GetAvailableSitesAsync()
        {
            var client = await _authService.GetAuthenticatedClientAsync();
            var sites = await client.Sites.GetAsync(
                config => config.QueryParameters.Filter = "siteCollection ne null"
            );

            var result = new List<SharePointSiteInfo>();
            if (sites?.Value != null)
            {
                foreach (var site in sites.Value)
                {
                    result.Add(new SharePointSiteInfo
                    {
                        Id = site.Id ?? string.Empty,
                        Name = site.Name ?? string.Empty,
                        DisplayName = site.DisplayName ?? site.Name ?? string.Empty,
                        WebUrl = site.WebUrl ?? string.Empty
                    });
                }
            }

            return result;
        }

        public async Task SetActiveSiteAsync(string siteId)
        {
            if (string.IsNullOrWhiteSpace(siteId))
                throw new ArgumentException("siteId ne peut pas être vide.", nameof(siteId));

            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                _activeSiteId = siteId.Trim();

                var model = new StoreModel
                {
                    SiteId = _activeSiteId,
                    SavedAtUtc = DateTime.UtcNow
                };

                var json = JsonSerializer.Serialize(model, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                Directory.CreateDirectory(Path.GetDirectoryName(_storePath)!);
                await File.WriteAllTextAsync(_storePath, json).ConfigureAwait(false);
            }
            finally
            {
                _lock.Release();
            }
        }

        public string GetActiveSiteId()
        {
            // Recharge si besoin (par sécurité si le champ est null)
            if (string.IsNullOrEmpty(_activeSiteId))
            {
                try { LoadFromDisk(); } catch { /* ignore */ }
            }

            return _activeSiteId ?? string.Empty;
        }

        public bool HasActiveSite()
        {
            if (!string.IsNullOrEmpty(_activeSiteId))
                return true;

            // Tentative de rechargement silencieux si la valeur n'est pas en mémoire
            try { LoadFromDisk(); } catch { /* ignore */ }

            return !string.IsNullOrEmpty(_activeSiteId);
        }

        // -------------------- Helpers persistance --------------------

        private void LoadFromDisk()
        {
            if (!File.Exists(_storePath))
            {
                _activeSiteId = null;
                return;
            }

            var json = File.ReadAllText(_storePath);
            if (string.IsNullOrWhiteSpace(json))
            {
                _activeSiteId = null;
                return;
            }

            var model = JsonSerializer.Deserialize<StoreModel>(json);
            _activeSiteId = string.IsNullOrWhiteSpace(model?.SiteId) ? null : model.SiteId.Trim();
        }
    }
}
