using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks; // AJOUTÉ
using System.Windows.Input;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class ConfigurationViewModel : ViewModelBase
    {
        private readonly ISharePointProvisioningService _provisioningService;
        private readonly IGraphAuthService _authService;

        private string _siteUrl = "https://grouperenault.sharepoint.com/sites/ShutdownMaintenance/";
        private string _statusMessage = "Prêt à exécuter le provisioning sur le site SharePoint.";
        private bool _isProvisioning;

        public string SiteUrl
        {
            get => _siteUrl;
            set => SetProperty(ref _siteUrl, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public bool IsProvisioning
        {
            get => _isProvisioning;
            set => SetProperty(ref _isProvisioning, value);
        }

        public ObservableCollection<string> LogEntries { get; } = new();

        public ICommand StartProvisioningCommand { get; }
        public ICommand TestConnectionCommand { get; }

        public ConfigurationViewModel(
            ISharePointProvisioningService provisioningService,
            IGraphAuthService authService)
        {
            _provisioningService = provisioningService;
            _authService = authService;

            StartProvisioningCommand = new SmartOffice365.UI.ViewModels.Base.RelayCommand(async () => await StartProvisioningAsync());
            TestConnectionCommand = new SmartOffice365.UI.ViewModels.Base.RelayCommand(async () => await TestConnectionAsync());
        }

        private async Task StartProvisioningAsync()
        {
            if (string.IsNullOrWhiteSpace(SiteUrl))
            {
                StatusMessage = "❌ Veuillez saisir une URL SharePoint valide.";
                return;
            }

            // On informe le service d'authentification de la nouvelle URL (qui va recalculer le Tenant ID si besoin)
            _authService.UpdateSharePointUrl(SiteUrl);

            IsProvisioning = true;
            LogEntries.Clear();
            LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] 🚀 Démarrage du provisioning des listes SharePoint sur {SiteUrl}...");

            // CORRECTION : On transmet l'URL saisie au service avant de lancer le provisioning
            _provisioningService.SetSiteId(SiteUrl);

            var progress = new Progress<string>(msg =>
            {
                LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] {msg}");
            });

            try
            {
                var result = await _provisioningService.ProvisionAllListsAsync(progress);
                if (result.Success)
                {
                    StatusMessage = "✅ Provisioning SharePoint effectué avec succès !";
                }
                else
                {
                    StatusMessage = "❌ Des erreurs sont survenues lors du provisioning.";
                    foreach (var err in result.Errors)
                        LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ❌ {err}");
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"❌ Erreur : {ex.Message}";
                LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ❌ Exception : {ex.Message}");
            }
            finally
            {
                IsProvisioning = false;
            }
        }

        private async Task TestConnectionAsync()
        {
            if (string.IsNullOrWhiteSpace(SiteUrl))
            {
                StatusMessage = "❌ Veuillez saisir une URL SharePoint valide.";
                return;
            }

            LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] 🔍 Test d'accès au site SharePoint {SiteUrl}...");
            try
            {
                // On informe le service d'authentification de la nouvelle URL (qui va recalculer le Tenant ID si besoin)
                _authService.UpdateSharePointUrl(SiteUrl);

                // CORRECTION : On transmet l'URL saisie au service avant de lancer le test
                _provisioningService.SetSiteId(SiteUrl);

                var connected = await _provisioningService.TestSiteConnectionAsync();
                if (connected)
                {
                    LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ✅ Connexion au site SharePoint établie avec succès !");
                    StatusMessage = "✅ Site SharePoint accessible.";
                }
                else
                {
                    LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ⚠️ Impossible de joindre le site SharePoint (vérifiez l'URL).");
                    StatusMessage = "⚠️ Impossible de joindre le site SharePoint.";
                }
            }
            catch (Exception ex)
            {
                LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ❌ Échec de connexion : {ex.Message}");
                StatusMessage = $"❌ Erreur d'accès : {ex.Message}";
            }
        }
    }
}
