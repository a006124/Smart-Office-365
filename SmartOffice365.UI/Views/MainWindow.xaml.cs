using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Threading;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.UI.ViewModels;

namespace SmartOffice365.UI.Views
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private readonly IGraphAuthService _authService;
        private readonly ISharePointSelectionService _selectionService;
        private readonly ISharePointDataService _dataService;
        private readonly MainViewModel _mainViewModel;
        
        private string _statusMessage = "Prêt";
        private string _userName = "Utilisateur";
        private bool _isConnected = false;
        private string _connectionStatus = "Déconnecté";
        private string _activeSiteName = string.Empty;
        private bool _hasActiveSite = false;
        private string _currentTime = DateTime.Now.ToString("HH:mm:ss");
        
        private int _totalOT = 0;
        private int _totalOTTermines = 0;
        private int _totalOTBloques = 0;
        private int _totalOTEnRetard = 0;

        public event PropertyChangedEventHandler? PropertyChanged;

        public MainWindow(
            IGraphAuthService authService,
            ISharePointSelectionService selectionService,
            ISharePointDataService dataService,
            MainViewModel mainViewModel)
        {
            InitializeComponent();
            
            _authService = authService;
            _selectionService = selectionService;
            _dataService = dataService;
            _mainViewModel = mainViewModel;
            
            DataContext = this;
            
            Loaded += OnLoaded;
            Closing += OnClosing;
            
            // Timer pour l'horloge
            var timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            timer.Tick += (s, e) => CurrentTime = DateTime.Now.ToString("HH:mm:ss");
            timer.Start();
            
            // Timer pour mettre à jour les KPIs
            var kpiTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(30)
            };
            kpiTimer.Tick += async (s, e) => await UpdateKpisAsync();
            kpiTimer.Start();
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                StatusMessage = "Connexion à Office 365...";
                
                // Vérifier l'authentification
                IsConnected = await _authService.IsAuthenticatedAsync();
                ConnectionStatus = IsConnected ? "Connecté" : "Déconnecté";
                
                if (IsConnected)
                {
                    UserName = await _authService.GetCurrentUserDisplayNameAsync();
                }
                
                // Vérifier si un site est déjà sélectionné
                HasActiveSite = _selectionService.HasActiveSite();
                if (HasActiveSite)
                {
                    var sites = await _selectionService.GetAvailableSitesAsync();
                    var activeSite = sites.FirstOrDefault(s => s.Id == _selectionService.GetActiveSiteId());
                    ActiveSiteName = activeSite?.DisplayName ?? "Site sélectionné";
                    await UpdateKpisAsync();
                }
                
                StatusMessage = "Prêt";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erreur : {ex.Message}";
            }
        }

        private async void OnClosing(object? sender, CancelEventArgs e)
        {
            // Nettoyage si nécessaire
            await Task.CompletedTask;
        }

        private async Task UpdateKpisAsync()
        {
            if (!HasActiveSite) return;
            
            try
            {
                var kpis = await _dataService.GetDashboardKpisAsync();
                TotalOT = kpis.TotalOT;
                TotalOTTermines = kpis.OTTermines;
                TotalOTBloques = kpis.OTBloques;
                TotalOTEnRetard = kpis.OTEnRetard;
            }
            catch (Exception ex)
            {
                StatusMessage = $"Erreur KPI : {ex.Message}";
            }
        }

        // Propriétés bindées avec notification
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        public string UserName
        {
            get => _userName;
            set { _userName = value; OnPropertyChanged(); }
        }

        public bool IsConnected
        {
            get => _isConnected;
            set { _isConnected = value; OnPropertyChanged(); }
        }

        public string ConnectionStatus
        {
            get => _connectionStatus;
            set { _connectionStatus = value; OnPropertyChanged(); }
        }

        public string ActiveSiteName
        {
            get => _activeSiteName;
            set { _activeSiteName = value; OnPropertyChanged(); }
        }

        public bool HasActiveSite
        {
            get => _hasActiveSite;
            set { _hasActiveSite = value; OnPropertyChanged(); }
        }

        public string CurrentTime
        {
            get => _currentTime;
            set { _currentTime = value; OnPropertyChanged(); }
        }

        public int TotalOT
        {
            get => _totalOT;
            set { _totalOT = value; OnPropertyChanged(); }
        }

        public int TotalOTTermines
        {
            get => _totalOTTermines;
            set { _totalOTTermines = value; OnPropertyChanged(); }
        }

        public int TotalOTBloques
        {
            get => _totalOTBloques;
            set { _totalOTBloques = value; OnPropertyChanged(); }
        }

        public int TotalOTEnRetard
        {
            get => _totalOTEnRetard;
            set { _totalOTEnRetard = value; OnPropertyChanged(); }
        }

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}