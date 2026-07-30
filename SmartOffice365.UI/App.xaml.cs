using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Services;
using SmartOffice365.UI.ViewModels;
using SmartOffice365.UI.Views;

namespace SmartOffice365.UI
{
    public partial class App : System.Windows.Application
    {
        private ServiceProvider? _serviceProvider;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var services = new ServiceCollection();
            ConfigureServices(services);

            _serviceProvider = services.BuildServiceProvider();

            // Récupérer les services
            var authService = _serviceProvider.GetRequiredService<IGraphAuthService>();
            var selectionService = _serviceProvider.GetRequiredService<ISharePointSelectionService>();
            var dataService = _serviceProvider.GetRequiredService<ISharePointDataService>();
            var mainViewModel = _serviceProvider.GetRequiredService<MainViewModel>();

            // Créer et afficher la fenêtre principale
            var mainWindow = new MainWindow(
                authService,
                selectionService,
                dataService,
                mainViewModel);

            // Passer les services à la vue de sélection de site
            var siteSelectionView = mainWindow.FindName("SiteSelectionControl") as SiteSelectionView;
            if (siteSelectionView != null)
            {
                // CORRECTION : On demande au conteneur de services de créer le ViewModel avec toutes ses dépendances !
                var siteSelectionViewModel = _serviceProvider.GetRequiredService<SiteSelectionViewModel>();
                siteSelectionView.SetViewModel(siteSelectionViewModel);
            }

            mainWindow.Show();
        }

        private void ConfigureServices(IServiceCollection services)
        {
            // Services de base
            services.AddSingleton<IGraphAuthService, GraphAuthService>();
            services.AddSingleton<ISharePointSelectionService, SharePointSelectionService>();
            services.AddSingleton<ISharePointProvisioningService, SharePointProvisioningService>();
            services.AddSingleton<ISharePointDataService, SharePointDataService>();
            services.AddTransient<PrerequisViewModel>();


            // ViewModels
            services.AddTransient<MainViewModel>();
            services.AddTransient<SiteSelectionViewModel>(); // CORRECTION : Enregistrement du ViewModel ici !
            services.AddTransient<DashboardViewModel>();
            services.AddTransient<ContactsViewModel>();
            services.AddTransient<OrdresDeTravailViewModel>();
            services.AddTransient<RessourcesViewModel>();
            services.AddTransient<ConfigurationViewModel>();

            // Services supplémentaires (si utilisés)
            services.AddSingleton<IOutlookReportService, OutlookReportService>();
            services.AddSingleton<ITeamsNotificationService>(sp =>
            {
                // 1. On récupère le service d'authentification depuis le conteneur
                var authService = sp.GetRequiredService<IGraphAuthService>();

                // 2. On instancie le service avec les 3 paramètres requis (ici des valeurs fictives pour le dev)
                string defaultTeamId = "00000000-0000-0000-0000-000000000000";
                string defaultChannelId = "19:association-id@thread.v2";

                return new TeamsNotificationService(authService, defaultTeamId, defaultChannelId);
            });
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _serviceProvider?.Dispose();
            base.OnExit(e);
        }
    }
}
