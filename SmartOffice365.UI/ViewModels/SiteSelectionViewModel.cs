// language: csharp
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models;

namespace SmartOffice365.UI.ViewModels
{
    /// <summary>
    /// ViewModel pour la vue de sélection de site SharePoint
    /// </summary>
    public partial class SiteSelectionViewModel : ObservableObject
    {
        private readonly IGraphAuthService _authService;
        private readonly ISharePointSelectionService _selectionService;

        [ObservableProperty]
        private bool _isLoading;

        [ObservableProperty]
        private SharePointSiteInfo? _selectedSite;

        [ObservableProperty]
        private string _statusMessage = string.Empty;

        /// <summary>
        /// Liste des sites SharePoint disponibles
        /// </summary>
        public ObservableCollection<SharePointSiteInfo> Sites { get; } = new();

        public SiteSelectionViewModel(IGraphAuthService authService, ISharePointSelectionService selectionService)
        {
            _authService = authService;
            _selectionService = selectionService;
        }

        /// <summary>
        /// Commande asynchrone pour charger les sites SharePoint (via le service de sélection)
        /// </summary>
        [RelayCommand]
        private async Task LoadSitesAsync()
        {
            if (IsLoading) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Chargement des sites…";
                Sites.Clear();

                // Récupération via le service (qui s'appuie lui-même sur Microsoft Graph)
                var list = await _selectionService.GetAvailableSitesAsync();

                // Tri facultatif par DisplayName
                foreach (var s in list.OrderBy(s => s.DisplayName))
                    Sites.Add(s);

                // Pré‑sélection du site actif persisté
                if (_selectionService.HasActiveSite())
                {
                    var activeId = _selectionService.GetActiveSiteId();
                    var match = Sites.FirstOrDefault(s => string.Equals(s.Id, activeId, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                    {
                        SelectedSite = match;
                        StatusMessage = $"Site actif restauré: {match.DisplayName}";
                    }
                    else
                    {
                        StatusMessage = "Aucun site actif trouvé dans la liste chargée.";
                    }
                }
                else
                {
                    StatusMessage = Sites.Count > 0
                        ? "Sélectionnez un site puis validez."
                        : "Aucun site retourné par Microsoft Graph.";
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Erreur lors du chargement des sites SharePoint : {ex.Message}",
                                "Erreur de chargement",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                StatusMessage = "Erreur lors du chargement des sites.";
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// Commande asynchrone pour valider la sélection du site (persistance via ISharePointSelectionService)
        /// </summary>
        [RelayCommand]
        private async Task SelectSiteAsync()
        {
            if (SelectedSite == null)
            {
                StatusMessage = "Veuillez sélectionner un site.";
                return;
            }

            try
            {
                await _selectionService.SetActiveSiteAsync(SelectedSite.Id);

                StatusMessage = $"Site '{SelectedSite.DisplayName}' sélectionné et mémorisé.";
                System.Windows.MessageBox.Show($"Site '{SelectedSite.DisplayName}' sélectionné avec succès !",
                                "Sélection validée",
                                MessageBoxButton.OK,
                                MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Erreur lors de la sélection du site : {ex.Message}",
                                "Erreur",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
                StatusMessage = "Erreur lors de la sélection du site.";
            }
        }
    }
}
