using System.Collections.ObjectModel;
using System.Windows.Input;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class PrerequisViewModel : ViewModelBase
    {
        private readonly ISharePointDataService _dataService;

        public ObservableCollection<PrerequsEntity> PrerequisList { get; } = new();

        public PrerequisViewModel(ISharePointDataService dataService)
        {
            _dataService = dataService;
            LoadMockData();
        }

        private void LoadMockData()
        {
            PrerequisList.Clear();
            PrerequisList.Add(new PrerequsEntity { Id=1, NumeroOT="40001234", Type="Consignation électrique", EstValide=true, DateValidation=DateTime.Now.AddDays(-2), Signataire="M. Martin", DateExpiration=DateTime.Now.AddDays(5), Commentaire="Cadenassage TGBT vannes 1 & 2" });
            PrerequisList.Add(new PrerequsEntity { Id=2, NumeroOT="40001234", Type="Permis de feu", EstValide=false, Signataire="", DateExpiration=DateTime.Now.AddDays(1), Commentaire="Attente mesure d'explosimétrie" });
            PrerequisList.Add(new PrerequsEntity { Id=3, NumeroOT="40001236", Type="ATEX", EstValide=true, DateValidation=DateTime.Now.AddDays(-1), Signataire="P. Lefebvre", DateExpiration=DateTime.Now.AddDays(10), Commentaire="Zonage 1 validé HSE" });
        }
    }
}
