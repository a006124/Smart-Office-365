using System.Collections.ObjectModel;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class ContactsViewModel : ViewModelBase
    {
        private readonly ISharePointDataService _dataService;

        public ObservableCollection<ContactEntity> ContactsList { get; } = new();

        public ContactsViewModel(ISharePointDataService dataService)
        {
            _dataService = dataService;
            LoadMockData();
        }

        private void LoadMockData()
        {
            ContactsList.Clear();
            ContactsList.Add(new ContactEntity { Id=1, Nom="Dupont", Prenom="Jean", Role="Chef d'Arrêt", Email="j.dupont@usine.com", Telephone="+33 6 12 34 56 78", CompteTeams="j.dupont@usine.com", CodeVendorSapLIFNR="VND-001", Entreprise="Usine Interne" });
            ContactsList.Add(new ContactEntity { Id=2, Nom="Martin", Prenom="Michel", Role="Chargé de Consignation", Email="m.martin@usine.com", Telephone="+33 6 98 76 54 32", CompteTeams="m.martin@usine.com", CodeVendorSapLIFNR="VND-001", Entreprise="Usine Interne" });
            ContactsList.Add(new ContactEntity { Id=3, Nom="Lefebvre", Prenom="Pierre", Role="Superviseur Prestataire", Email="p.lefebvre@endel.fr", Telephone="+33 6 11 22 33 44", CompteTeams="p.lefebvre@endel.fr", CodeVendorSapLIFNR="LIF-88392", Entreprise="ENDEL ENGIE" });
        }
    }
}
