// language: csharp
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Services; // GraphAuthService, SharePointSelectionService, etc.
using SmartOffice365.UI.ViewModels;
using SmartOffice365.UI.Views;

namespace SmartOffice365.UI
{
    public partial class App : System.Windows.Application
    {
        private ServiceProvider? _serviceProvider;

        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var services = new ServiceCollection();
            ConfigureServices(services);

            _serviceProvider = services.BuildServiceProvider();

            // Sanity check: IGraphAuthService DOIT être résolvable ici
            try
            {
                var test = _serviceProvider.GetRequiredService<IGraphAuthService>();
                Debug.WriteLine("[DI] IGraphAuthService OK: " + test.GetType().FullName);
            }
            catch
            {
                System.Windows.MessageBox.Show("IGraphAuthService introuvable dans le conteneur DI. Vérifiez ConfigureServices.", "DI Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
                return;
            }

            var authService = _serviceProvider.GetRequiredService<IGraphAuthService>();
            var selectionService = _serviceProvider.GetRequiredService<ISharePointSelectionService>();
            var dataService = _serviceProvider.GetRequiredService<ISharePointDataService>();
            var provisioningService = _serviceProvider.GetRequiredService<ISharePointProvisioningService>();
            var mainViewModel = _serviceProvider.GetRequiredService<MainViewModel>(); // Singleton

            var mainWindow = new MainWindow(authService, selectionService, dataService, mainViewModel);

            // Injecte le VM de sélection via DI
            if (mainWindow.FindName("SiteSelectionControl") is SiteSelectionView siteSelectionView)
            {
                var siteSelectionViewModel = _serviceProvider.GetRequiredService<SiteSelectionViewModel>();
                siteSelectionView.SetViewModel(siteSelectionViewModel);
            }

            mainWindow.Show();

            // Statut immédiat
            mainViewModel.StatusMessage = "Vérification de la connexion SharePoint…";

            if (selectionService.HasActiveSite())
            {
                provisioningService.SetSiteId(selectionService.GetActiveSiteId());
                var ok = await provisioningService.TestSiteConnectionAsync();
                mainViewModel.StatusMessage = ok
                    ? "✓ Connexion SharePoint (site actif) OK"
                    : "⚠️ Connexion SharePoint (site actif) échouée";
            }
            else
            {
                mainViewModel.StatusMessage = "ℹ Aucun site SharePoint sélectionné. Ouvrez 'Sélection du site', choisissez un site puis cliquez 'Sélectionner'.";
            }
        }

        private void ConfigureServices(IServiceCollection services)
        {
            // Services Core (SINGLETONS)
            services.AddSingleton<IGraphAuthService, GraphAuthService>();                  // <- ENREGISTRÉ ICI
            services.AddSingleton<ISharePointService, SharePointService>();
            services.AddSingleton<ISharePointSelectionService, SharePointSelectionService>();
            services.AddSingleton<ISharePointProvisioningService, SharePointProvisioningService>();
            services.AddSingleton<ISharePointDataService, SharePointService>();

            // ViewModels
            services.AddSingleton<MainViewModel>();         // Singleton pour que la StatusBar se mette bien à jour
            services.AddTransient<SiteSelectionViewModel>();
            services.AddTransient<DashboardViewModel>();
            services.AddTransient<ArretsViewModel>();
            services.AddTransient<ContactsViewModel>();
            services.AddTransient<OrdresDeTravailViewModel>();
            services.AddTransient<RessourcesViewModel>();
            services.AddTransient<ConfigurationViewModel>();
            services.AddTransient<PrerequisViewModel>();

            // Services supplémentaires
            services.AddSingleton<IOutlookReportService, OutlookReportService>();
            services.AddSingleton<ITeamsNotificationService>(sp =>
            {
                var auth = sp.GetRequiredService<IGraphAuthService>(); // <- résolu AU MOMENT de la création
                string defaultTeamId = "00000000-0000-0000-0000-000000000000";
                string defaultChannelId = "19:association-id@thread.v2";
                return new TeamsNotificationService(auth, defaultTeamId, defaultChannelId);
            });
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _serviceProvider?.Dispose();
            base.OnExit(e);
        }
    }
}
