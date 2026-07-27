using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Services;
using SmartOffice365.UI.ViewModels;
using SmartOffice365.UI.Views;

namespace SmartOffice365.UI
{
    public partial class App : Application
    {
        private IHost? _host;

        public App()
        {
            _host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) =>
                {
                    // Core Services (Authentification sans ClientId via connexion Office 365 directe)
                    services.AddSingleton<IGraphAuthService, GraphAuthService>();

                    services.AddSingleton<ISharePointProvisioningService>(sp =>
                        new SharePointProvisioningService(sp.GetRequiredService<IGraphAuthService>(), "root"));
                    services.AddSingleton<ISharePointDataService>(sp =>
                        new SharePointDataService(sp.GetRequiredService<IGraphAuthService>(), "root"));
                    services.AddSingleton<ITeamsNotificationService>(sp =>
                        new TeamsNotificationService(sp.GetRequiredService<IGraphAuthService>(), "TEAM-ID", "CHANNEL-ID"));
                    services.AddSingleton<IOutlookReportService, OutlookReportService>();

                    // ViewModels
                    services.AddSingleton<MainViewModel>();
                    services.AddTransient<DashboardViewModel>();
                    services.AddTransient<OrdresDeTravailViewModel>();
                    services.AddTransient<PrerequisViewModel>();
                    services.AddTransient<RessourcesViewModel>();
                    services.AddTransient<ContactsViewModel>();
                    services.AddTransient<ConfigurationViewModel>();

                    // Views
                    services.AddSingleton<MainWindow>();
                })
                .Build();
        }

        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            await _host!.StartAsync();

            var mainWindow = _host.Services.GetRequiredService<MainWindow>();
            mainWindow.DataContext = _host.Services.GetRequiredService<MainViewModel>();
            mainWindow.Show();
        }

        protected override async void OnExit(ExitEventArgs e)
        {
            if (_host != null)
            {
                await _host.StopAsync();
                _host.Dispose();
            }
            base.OnExit(e);
        }
    }
}
