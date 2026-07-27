using System.Collections.ObjectModel;
using System.Windows.Input;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class DashboardViewModel : ViewModelBase
    {
        private readonly ISharePointDataService _dataService;
        private readonly ITeamsNotificationService _teamsService;
        private readonly IOutlookReportService _outlookService;

        private DashboardKpis _kpis = new();
        private bool _isLoading;

        public DashboardKpis Kpis
        {
            get => _kpis;
            set => SetProperty(ref _kpis, value);
        }

        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }

        public ObservableCollection<OrdreDeTravailEntity> OtAfaire { get; } = new();
        public ObservableCollection<OrdreDeTravailEntity> OtEnCours { get; } = new();
        public ObservableCollection<OrdreDeTravailEntity> OtBloques { get; } = new();
        public ObservableCollection<OrdreDeTravailEntity> OtTermines { get; } = new();

        public ICommand RefreshCommand { get; }
        public ICommand SendDailyReportCommand { get; }

        public DashboardViewModel(
            ISharePointDataService dataService,
            ITeamsNotificationService teamsService,
            IOutlookReportService outlookService)
        {
            _dataService = dataService;
            _teamsService = teamsService;
            _outlookService = outlookService;

            RefreshCommand = new AsyncRelayCommand(LoadDataAsync);
            SendDailyReportCommand = new AsyncRelayCommand(SendDailyReportAsync);

            // Charger les données initiales ou données de démo si hors-ligne
            LoadMockData();
        }

        public async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true;
                Kpis = await _dataService.GetDashboardKpisAsync();

                var ots = await _dataService.GetOrdresDeTravailAsync();
                PopulateKanban(ots);
            }
            catch
            {
                // En cas d'erreur de connexion Graph API, fallback sur données de démo
                LoadMockData();
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task SendDailyReportAsync()
        {
            try
            {
                var otsEnRetard = OtBloques.Concat(OtEnCours).Where(o => o.DateFinPrevue < DateTime.Now).ToList();
                await _outlookService.SendDailyReportAsync(
                    new[] { "responsable-shutdown@entreprise.com" }, Kpis, otsEnRetard);
                await _teamsService.SendDailyProgressCardAsync(Kpis);
            }
            catch { }
        }

        private void PopulateKanban(List<OrdreDeTravailEntity> ots)
        {
            OtAfaire.Clear();
            OtEnCours.Clear();
            OtBloques.Clear();
            OtTermines.Clear();

            foreach (var ot in ots)
            {
                switch (ot.StatutShutdown)
                {
                    case "Bloqué": OtBloques.Add(ot); break;
                    case "En cours": OtEnCours.Add(ot); break;
                    case "Terminé": OtTermines.Add(ot); break;
                    default: OtAfaire.Add(ot); break;
                }
            }
        }

        private void LoadMockData()
        {
            Kpis = new DashboardKpis
            {
                TotalOT = 24,
                OTTermines = 10,
                OTEnCours = 8,
                OTBloques = 3,
                OTEnRetard = 2,
                AvancementGlobal = 58.5,
                TotalContacts = 42,
                TotalAffaires = 3
            };

            var mockOts = new List<OrdreDeTravailEntity>
            {
                new() { Id=1, NumeroOT_Aufnr="40001234", Titre="Remplacement Vane V-102 Unité Craquage", StatutShutdown="Bloqué", Priorite="Critique", Avancement=20, Responsable="J. Dupont", MotifsBlockage="Pièce de rechange en attente douane", NumeroEquipement_EQUNR="EQ-9921", PosteTechnique_TPLNR="U100-CR-01" },
                new() { Id=2, NumeroOT_Aufnr="40001235", Titre="Consignation Électrique TGBT-02", StatutShutdown="Terminé", Priorite="Haute", Avancement=100, Responsable="M. Martin", NumeroEquipement_EQUNR="EQ-1002", PosteTechnique_TPLNR="U100-EL-02" },
                new() { Id=3, NumeroOT_Aufnr="40001236", Titre="Épreuve Hydraulique Réacteur R-201", StatutShutdown="En cours", Priorite="Critique", Avancement=65, Responsable="P. Lefebvre", NumeroEquipement_EQUNR="EQ-3301", PosteTechnique_TPLNR="U200-RX-01" },
                new() { Id=4, NumeroOT_Aufnr="40001237", Titre="Contrôle Ultrasons Soudures Tuyauterie", StatutShutdown="En cours", Priorite="Normale", Avancement=40, Responsable="A. Bernard", NumeroEquipement_EQUNR="EQ-4410", PosteTechnique_TPLNR="U100-PIP-05" },
                new() { Id=5, NumeroOT_Aufnr="40001238", Titre="Nettoyage Chimique Échangeur E-104", StatutShutdown="À Faire", Priorite="Normale", Avancement=0, Responsable="C. Thomas", NumeroEquipement_EQUNR="EQ-2204", PosteTechnique_TPLNR="U100-EX-04" },
                new() { Id=6, NumeroOT_Aufnr="40001239", Titre="Remplacement Garniture Pompe P-302B", StatutShutdown="Bloqué", Priorite="Haute", Avancement=10, Responsable="J. Dupont", MotifsBlockage="Habilitation ATEX prestataire expirée", NumeroEquipement_EQUNR="EQ-5002", PosteTechnique_TPLNR="U300-P-02" },
            };

            PopulateKanban(mockOts);
        }
    }
}
