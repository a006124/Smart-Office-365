using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows.Input;
using SmartOffice365.UI.ViewModels.Base;

namespace SmartOffice365.UI.ViewModels
{
    public class ArretsViewModel : ViewModelBase
    {
        private readonly ISharePointService _sharePointService; // AJOUTÉ
        private ObservableCollection<ArretModel> _arrets;
        private ArretModel _selectedArret;
        private ArretModel _editingArret;
        private bool _isEditMode;
        private bool _isLoading;

        public ObservableCollection<ArretModel> Arrets
        {
            get => _arrets;
            set { _arrets = value; OnPropertyChanged(); }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set { _isLoading = value; OnPropertyChanged(); }
        }

        // ... Conservez les autres propriétés (SelectedArret, EditingArret, IsEditMode) ...

        public ICommand NouvelArretCommand { get; }
        public ICommand EnregistrerCommand { get; }
        public ICommand SupprimerCommand { get; }
        public ICommand AnnulerCommand { get; }

        // Le constructeur accepte désormais le service SharePoint
        public ArretsViewModel(ISharePointService sharePointService)
        {
            _sharePointService = sharePointService;
            Arrets = new ObservableCollection<ArretModel>();

            NouvelArretCommand = new RelayCommand(ExecuteNouvelArret);
            EnregistrerCommand = new RelayCommand(async (obj) => await ExecuteEnregistrerAsync(), CanExecuteEnregistrer);
            SupprimerCommand = new RelayCommand(async (obj) => await ExecuteSupprimerAsync(), CanExecuteSelection);
            AnnulerCommand = new RelayCommand(ExecuteAnnuler);

            // Charger les données depuis SharePoint au démarrage
            _ = LoadDataFromSharePointAsync();
        }

        private async Task LoadDataFromSharePointAsync()
        {
            try
            {
                IsLoading = true;
                var data = await _sharePointService.GetArretsAsync();
                Arrets = new ObservableCollection<ArretModel>(data);
            }
            catch (Exception ex)
            {
                // Gérer l'erreur (ex: afficher un message dans la barre d'état)
                System.Diagnostics.Debug.WriteLine($"Erreur de chargement SharePoint : {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ExecuteNouvelArret(object obj)
        {
            SelectedArret = null;
            EditingArret = new ArretModel
            {
                DateDebut = DateTime.Now.AddDays(7),
                DateFin = DateTime.Now.AddDays(14),
                Statut = "En préparation"
            };
            IsEditMode = true;
        }

        private async Task ExecuteEnregistrerAsync()
        {
            try
            {
                IsLoading = true;
                if (EditingArret.Id == 0) // Création
                {
                    var nouvelArret = await _sharePointService.CreateArretAsync(EditingArret);
                    Arrets.Add(nouvelArret);
                }
                else // Modification
                {
                    await _sharePointService.UpdateArretAsync(EditingArret);
                    // Mettre à jour l'élément dans la liste locale
                    for (int i = 0; i < Arrets.Count; i++)
                    {
                        if (Arrets[i].Id == EditingArret.Id)
                        {
                            Arrets[i] = EditingArret;
                            break;
                        }
                    }
                }
                SelectedArret = null;
                IsEditMode = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur d'enregistrement : {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task ExecuteSupprimerAsync()
        {
            if (SelectedArret != null)
            {
                try
                {
                    IsLoading = true;
                    await _sharePointService.DeleteArretAsync(SelectedArret.Id);
                    Arrets.Remove(SelectedArret);
                    SelectedArret = null;
                    IsEditMode = false;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Erreur de suppression : {ex.Message}");
                }
                finally
                {
                    IsLoading = false;
                }
            }
        }


private bool CanExecuteEnregistrer(object obj)
        {
            return EditingArret != null && !string.IsNullOrWhiteSpace(EditingArret.Titre);
        }

        private void ExecuteSupprimer(object obj)
        {
            if (SelectedArret != null)
            {
                Arrets.Remove(SelectedArret);
                SelectedArret = null;
                IsEditMode = false;
            }
        }

        private bool CanExecuteSelection(object obj)
        {
            return SelectedArret != null;
        }

        private void ExecuteAnnuler(object obj)
        {
            SelectedArret = null;
            IsEditMode = false;
        }
    }

    // Modèle de données local (conserve son implémentation légère)
    public class ArretModel : System.ComponentModel.INotifyPropertyChanged
    {
        private int _id;
        private string _titre;
        private DateTime _dateDebut;
        private DateTime _dateFin;
        private string _statut;
        private string _description;
        private string _jalonsPreparation;

        public int Id { get => _id; set { _id = value; OnPropertyChanged(); } }
        public string Titre { get => _titre; set { _titre = value; OnPropertyChanged(); } }
        public DateTime DateDebut { get => _dateDebut; set { _dateDebut = value; OnPropertyChanged(); } }
        public DateTime DateFin { get => _dateFin; set { _dateFin = value; OnPropertyChanged(); } }
        public string Statut { get => _statut; set { _statut = value; OnPropertyChanged(); } }
        public string Description { get => _description; set { _description = value; OnPropertyChanged(); } }
        public string JalonsPreparation { get => _jalonsPreparation; set { _jalonsPreparation = value; OnPropertyChanged(); } }

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
        }
    }
}
