using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Services;
using SmartOffice365.UI.ViewModels;
using SmartOffice365.UI.Views;

namespace SmartOffice365.UI
{
    public partial class App : Application
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
            var provisioningService = _serviceProvider.GetRequiredService<ISharePointProvisioningService>();
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
                var siteSelectionViewModel = new SiteSelectionViewModel(
                    selectionService,
                    provisioningService);
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
            
            // ViewModels
            services.AddTransient<MainViewModel>();
            services.AddTransient<DashboardViewModel>();
            services.AddTransient<ContactsViewModel>();
            services.AddTransient<OrdresDeTravailViewModel>();
            services.AddTransient<RessourcesViewModel>();
            services.AddTransient<ConfigurationViewModel>();
            
            // Services supplémentaires (si utilisés)
            services.AddSingleton<IOutlookReportService, OutlookReportService>();
            services.AddSingleton<ITeamsNotificationService, TeamsNotificationService>();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _serviceProvider?.Dispose();
            base.OnExit(e);
        }
    }
}