using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks; // AJOUTÉ
using System.Windows.Input;
using CommunityToolkit.Mvvm.Input;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private readonly IGraphAuthService _authService; // AJOUTÉ
        private ViewModelBase _currentView;
        private string _currentTitle = "Tableau de Bord";
        private string _searchQuery = string.Empty;
        private string _userDisplayName = "Connexion en cours..."; // Valeur d'attente
        private bool _isSidebarExpanded = true;

        public ViewModelBase CurrentView
        {
            get => _currentView;
            set => SetProperty(ref _currentView, value);
        }

        public string CurrentTitle
        {
            get => _currentTitle;
            set => SetProperty(ref _currentTitle, value);
        }

        public string SearchQuery
        {
            get => _searchQuery;
            set => SetProperty(ref _searchQuery, value);
        }

        public string UserDisplayName
        {
            get => _userDisplayName;
            set => SetProperty(ref _userDisplayName, value);
        }

        public bool IsSidebarExpanded
        {
            get => _isSidebarExpanded;
            set => SetProperty(ref _isSidebarExpanded, value);
        }

        private string _statusMessage = string.Empty;
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }
        
        // ViewModels enfants
        public DashboardViewModel DashboardVM { get; }
        public ArretsViewModel ArretsVM { get; }
        public OrdresDeTravailViewModel OrdresDeTravailVM { get; }
        public PrerequisViewModel PrerequisVM { get; }
        public RessourcesViewModel RessourcesVM { get; }
        public ContactsViewModel ContactsVM { get; }
        public ConfigurationViewModel ConfigurationVM { get; }

        // Commandes de navigation
        public ICommand NavigateCommand { get; }
        public ICommand ToggleSidebarCommand { get; }

        // Le constructeur accepte désormais IGraphAuthService en premier paramètre
        public MainViewModel(
            IGraphAuthService authService, // AJOUTÉ
            DashboardViewModel dashboardVM,
            ArretsViewModel arretsVM,
            OrdresDeTravailViewModel ordresDeTravailVM,
            PrerequisViewModel prerequisVM,
            RessourcesViewModel ressourcesVM,
            ContactsViewModel contactsVM,
            ConfigurationViewModel configurationVM)
        {
            _authService = authService; // AJOUTÉ

            DashboardVM = dashboardVM;
            ArretsVM = arretsVM;
            OrdresDeTravailVM = ordresDeTravailVM;
            PrerequisVM = prerequisVM;
            RessourcesVM = ressourcesVM;
            ContactsVM = contactsVM;
            ConfigurationVM = configurationVM;
            {
                // init existante
                // Message d’accueil par défaut (optionnel)
                StatusMessage = "Application prête.";
            }

            _currentView = DashboardVM;

            NavigateCommand = new SmartOffice365.UI.ViewModels.Base.RelayCommand(param => Navigate(param?.ToString() ?? "Dashboard"));
            ToggleSidebarCommand = new SmartOffice365.UI.ViewModels.Base.RelayCommand(() => IsSidebarExpanded = !IsSidebarExpanded);

            // Lance la récupération du nom d'utilisateur en arrière-plan sans bloquer l'affichage
            _ = LoadUserDisplayNameAsync(); // AJOUTÉ
        }

        /// <summary>
        /// Récupère le nom de l'utilisateur connecté depuis Microsoft Graph de manière asynchrone
        /// </summary>
        private async Task LoadUserDisplayNameAsync() // AJOUTÉ
        {
            try
            {
                // Appel silencieux à Microsoft Graph (utilise votre session Windows/Office active)
                UserDisplayName = await _authService.GetCurrentUserDisplayNameAsync();
            }
            catch
            {
                // En cas d'échec, on applique un nom par défaut
                UserDisplayName = "Session Office 365 Active";
            }
        }

        public void Navigate(string destination)
        {
            switch (destination)
            {
                case "Dashboard":
                    CurrentView = DashboardVM;
                    CurrentTitle = "Cockpit de Pilotage";
                    break;
                case "Arrets":
                    CurrentView = ArretsVM;
                    CurrentTitle = "Gestion des Arrêts";
                    break;
                case "OrdresDeTravail":
                    CurrentView = OrdresDeTravailVM;
                    CurrentTitle = "Ordres de Travail & SAP PM";
                    break;
                case "Prerequis":
                    CurrentView = PrerequisVM;
                    CurrentTitle = "Prérequis & Consignations";
                    break;
                case "Ressources":
                    CurrentView = RessourcesVM;
                    CurrentTitle = "Ressources & Prestataires";
                    break;
                case "Contacts":
                    CurrentView = ContactsVM;
                    CurrentTitle = "Contacts & Habilitations";
                    break;
                case "Configuration":
                    CurrentView = ConfigurationVM;
                    CurrentTitle = "Administration & Provisioning";
                    break;
            }
        }
    }
}
