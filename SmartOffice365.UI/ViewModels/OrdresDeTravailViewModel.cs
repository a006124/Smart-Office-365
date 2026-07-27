using System.Collections.ObjectModel;
using System.Windows.Input;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class OrdresDeTravailViewModel : ViewModelBase
    {
        private readonly ISharePointDataService _dataService;
        private readonly ITeamsNotificationService _teamsService;

        private OrdreDeTravailEntity? _selectedOT;
        private string _searchText = string.Empty;
        private string _selectedPrioriteFilter = "Tous";
        private string _selectedStatutFilter = "Tous";
        private bool _isEditing;

        public ObservableCollection<OrdreDeTravailEntity> AllOrdresDeTravail { get; } = new();
        public ObservableCollection<OrdreDeTravailEntity> FilteredOrdresDeTravail { get; } = new();

        public OrdreDeTravailEntity? SelectedOT
        {
            get => _selectedOT;
            set
            {
                if (SetProperty(ref _selectedOT, value))
                {
                    IsEditing = _selectedOT != null;
                }
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value))
                    ApplyFilter();
            }
        }

        public string SelectedPrioriteFilter
        {
            get => _selectedPrioriteFilter;
            set
            {
                if (SetProperty(ref _selectedPrioriteFilter, value))
                    ApplyFilter();
            }
        }

        public string SelectedStatutFilter
        {
            get => _selectedStatutFilter;
            set
            {
                if (SetProperty(ref _selectedStatutFilter, value))
                    ApplyFilter();
            }
        }

        public bool IsEditing
        {
            get => _isEditing;
            set => SetProperty(ref _isEditing, value);
        }

        public List<string> Statuts { get; } = new() { "À Faire", "En cours", "Bloqué", "Terminé" };
        public List<string> Priorites { get; } = new() { "Critique", "Haute", "Normale", "Basse" };

        public ICommand SaveUpdateCommand { get; }
        public ICommand SendBlockedAlertCommand { get; }
        public ICommand RefreshCommand { get; }

        public OrdresDeTravailViewModel(
            ISharePointDataService dataService,
            ITeamsNotificationService teamsService)
        {
            _dataService = dataService;
            _teamsService = teamsService;

            SaveUpdateCommand = new AsyncRelayCommand(SaveUpdateAsync);
            SendBlockedAlertCommand = new AsyncRelayCommand(SendBlockedAlertAsync);
            RefreshCommand = new AsyncRelayCommand(LoadDataAsync);

            LoadMockData();
        }

        public async Task LoadDataAsync()
        {
            try
            {
                var ots = await _dataService.GetOrdresDeTravailAsync();
                AllOrdresDeTravail.Clear();
                foreach (var ot in ots) AllOrdresDeTravail.Add(ot);
                ApplyFilter();
            }
            catch { LoadMockData(); }
        }

        private async Task SaveUpdateAsync()
        {
            if (SelectedOT == null) return;

            try
            {
                await _dataService.UpdateAvancementAsync(
                    SelectedOT.Id, SelectedOT.Avancement, SelectedOT.StatutShutdown);

                if (SelectedOT.StatutShutdown == "Bloqué")
                {
                    await _teamsService.SendOTBlockedAlertAsync(SelectedOT);
                }
            }
            catch { }
        }

        private async Task SendBlockedAlertAsync()
        {
            if (SelectedOT != null && SelectedOT.StatutShutdown == "Bloqué")
            {
                await _teamsService.SendOTBlockedAlertAsync(SelectedOT);
            }
        }

        private void ApplyFilter()
        {
            FilteredOrdresDeTravail.Clear();
            var query = AllOrdresDeTravail.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                var s = SearchText.ToLower();
                query = query.Where(o =>
                    o.NumeroOT_Aufnr.ToLower().Contains(s) ||
                    o.NumeroEquipement_EQUNR.ToLower().Contains(s) ||
                    o.PosteTechnique_TPLNR.ToLower().Contains(s) ||
                    o.Titre.ToLower().Contains(s) ||
                    o.Responsable.ToLower().Contains(s));
            }

            if (SelectedPrioriteFilter != "Tous")
                query = query.Where(o => o.Priorite == SelectedPrioriteFilter);

            if (SelectedStatutFilter != "Tous")
                query = query.Where(o => o.StatutShutdown == SelectedStatutFilter);

            foreach (var item in query) FilteredOrdresDeTravail.Add(item);
        }

        private void LoadMockData()
        {
            AllOrdresDeTravail.Clear();
            var mock = new List<OrdreDeTravailEntity>
            {
                new() { Id=1, NumeroOT_Aufnr="40001234", Titre="Remplacement Vane V-102 Unité Craquage", StatutShutdown="Bloqué", Priorite="Critique", Avancement=20, Responsable="J. Dupont", MotifsBlockage="Pièce de rechange en attente douane", NumeroEquipement_EQUNR="EQ-9921", PosteTechnique_TPLNR="U100-CR-01", PosteTravail_ARBPL="MECANIK" },
                new() { Id=2, NumeroOT_Aufnr="40001235", Titre="Consignation Électrique TGBT-02", StatutShutdown="Terminé", Priorite="Haute", Avancement=100, Responsable="M. Martin", NumeroEquipement_EQUNR="EQ-1002", PosteTechnique_TPLNR="U100-EL-02", PosteTravail_ELECTR" },
                new() { Id=3, NumeroOT_Aufnr="40001236", Titre="Épreuve Hydraulique Réacteur R-201", StatutShutdown="En cours", Priorite="Critique", Avancement=65, Responsable="P. Lefebvre", NumeroEquipement_EQUNR="EQ-3301", PosteTechnique_TPLNR="U200-RX-01", PosteTravail_INSPECTION" },
                new() { Id=4, NumeroOT_Aufnr="40001237", Titre="Contrôle Ultrasons Soudures Tuyauterie", StatutShutdown="En cours", Priorite="Normale", Avancement=40, Responsable="A. Bernard", NumeroEquipement_EQUNR="EQ-4410", PosteTechnique_TPLNR="U100-PIP-05", PosteTravail_TUYAUTERIE" },
                new() { Id=5, NumeroOT_Aufnr="40001238", Titre="Nettoyage Chimique Échangeur E-104", StatutShutdown="À Faire", Priorite="Normale", Avancement=0, Responsable="C. Thomas", NumeroEquipement_EQUNR="EQ-2204", PosteTechnique_TPLNR="U100-EX-04", PosteTravail_NETTOYAGE" },
                new() { Id=6, NumeroOT_Aufnr="40001239", Titre="Remplacement Garniture Pompe P-302B", StatutShutdown="Bloqué", Priorite="Haute", Avancement=10, Responsable="J. Dupont", MotifsBlockage="Habilitation ATEX prestataire expirée", NumeroEquipement_EQUNR="EQ-5002", PosteTechnique_TPLNR="U300-P-02", PosteTravail_MECANIK" },
            };

            foreach (var ot in mock) AllOrdresDeTravail.Add(ot);
            ApplyFilter();
        }
    }
}
