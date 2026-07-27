using System.Windows.Input;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private ViewModelBase _currentView;
        private string _currentTitle = "Tableau de Bord";
        private string _searchQuery = string.Empty;
        private string _userDisplayName = "Chargement...";
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

        // ViewModels enfants
        public DashboardViewModel DashboardVM { get; }
        public OrdresDeTravailViewModel OrdresDeTravailVM { get; }
        public PrerequisViewModel PrerequisVM { get; }
        public RessourcesViewModel RessourcesVM { get; }
        public ContactsViewModel ContactsVM { get; }
        public ConfigurationViewModel ConfigurationVM { get; }

        // Commandes de navigation
        public ICommand NavigateCommand { get; }
        public ICommand ToggleSidebarCommand { get; }

        public MainViewModel(
            DashboardViewModel dashboardVM,
            OrdresDeTravailViewModel ordresDeTravailVM,
            PrerequisViewModel prerequisVM,
            RessourcesViewModel ressourcesVM,
            ContactsViewModel contactsVM,
            ConfigurationViewModel configurationVM)
        {
            DashboardVM = dashboardVM;
            OrdresDeTravailVM = ordresDeTravailVM;
            PrerequisVM = prerequisVM;
            RessourcesVM = ressourcesVM;
            ContactsVM = contactsVM;
            ConfigurationVM = configurationVM;

            _currentView = DashboardVM;

            NavigateCommand = new RelayCommand(param => Navigate(param?.ToString() ?? "Dashboard"));
            ToggleSidebarCommand = new RelayCommand(() => IsSidebarExpanded = !IsSidebarExpanded);
        }

        public void Navigate(string destination)
        {
            switch (destination)
            {
                case "Dashboard":
                    CurrentView = DashboardVM;
                    CurrentTitle = "Cockpit de Pilotage";
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
