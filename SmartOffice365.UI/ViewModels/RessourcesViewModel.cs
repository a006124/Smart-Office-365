using System.Collections.ObjectModel;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class RessourcesViewModel : ViewModelBase
    {
        private readonly ISharePointDataService _dataService;

        public ObservableCollection<RessourceEntity> RessourcesList { get; } = new();

        public RessourcesViewModel(ISharePointDataService dataService)
        {
            _dataService = dataService;
            LoadMockData();
        }

        private void LoadMockData()
        {
            RessourcesList.Clear();
            RessourcesList.Add(new RessourceEntity { Id=1, NumeroOT="40001234", Type="Main d'œuvre", EntreprisePrestataire="ENDEL ENGIE", Description="Équipe Tuyauterie HT", QuantitePrevue=4, QuantiteReelle=4, Unite="h", EstDisponible=true });
            RessourcesList.Add(new RessourceEntity { Id=2, NumeroOT="40001234", Type="Outillage spécial", EntreprisePrestataire="FAGOR", Description="Clé dynamométrique 2000 Nm", QuantitePrevue=1, QuantiteReelle=0, Unite="pcs", EstDisponible=false, Commentaire="En réétalonnage" });
            RessourcesList.Add(new RessourceEntity { Id=3, NumeroOT="40001236", Type="Engin", EntreprisePrestataire="MEDIACAKO", Description="Grue 50T Télescopique", QuantitePrevue=1, QuantiteReelle=1, Unite="pcs", EstDisponible=true });
        }
    }
}
