using System.Collections.ObjectModel;
using System.Windows.Input;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class ConfigurationViewModel : ViewModelBase
    {
        private readonly ISharePointProvisioningService _provisioningService;
        private readonly IGraphAuthService _authService;

        private string _siteUrl = "https://votre-tenant.sharepoint.com/sites/SmartMaintenance";
        private string _statusMessage = "Prêt à connecter votre compte Office 365.";
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

        public ICommand LoginOffice365Command { get; }
        public ICommand StartProvisioningCommand { get; }
        public ICommand TestConnectionCommand { get; }

        public ConfigurationViewModel(
            ISharePointProvisioningService provisioningService,
            IGraphAuthService authService)
        {
            _provisioningService = provisioningService;
            _authService = authService;

            LoginOffice365Command = new AsyncRelayCommand(LoginOffice365Async);
            StartProvisioningCommand = new AsyncRelayCommand(StartProvisioningAsync);
            TestConnectionCommand = new AsyncRelayCommand(TestConnectionAsync);
        }

        private async Task LoginOffice365Async()
        {
            LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] 🔐 Ouverture de la fenêtre de connexion Office 365...");
            StatusMessage = "Connexion Office 365 en cours dans le navigateur...";

            var success = await _authService.SignInAsync();
            if (success)
            {
                var name = await _authService.GetCurrentUserDisplayNameAsync();
                StatusMessage = $"✅ Connecté en tant que : {name}";
                LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ✅ Authentification Office 365 réussie pour {name}");
            }
            else
            {
                StatusMessage = "❌ Connexion annulée ou échouée.";
                LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ❌ Échec de l'authentification Office 365.");
            }
        }

        private async Task StartProvisioningAsync()
        {
            IsProvisioning = true;
            LogEntries.Clear();
            LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] 🚀 Démarrage du provisioning automatique des listes SharePoint...");

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
            LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] 🔍 Test de connexion au site SharePoint...");
            try
            {
                var connected = await _provisioningService.TestSiteConnectionAsync();
                if (connected)
                {
                    LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ✅ Connexion au site SharePoint établie !");
                }
                else
                {
                    LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ⚠️ Impossible d'accéder à SharePoint. Connectez votre compte Office 365.");
                }
            }
            catch (Exception ex)
            {
                LogEntries.Add($"[{DateTime.Now:HH:mm:ss}] ❌ Échec de connexion : {ex.Message}");
            }
        }
    }
}
