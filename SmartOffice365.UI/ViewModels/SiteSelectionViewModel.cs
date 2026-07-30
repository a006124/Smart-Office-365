using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SmartOffice365.Core.Interfaces;
using SmartOffice365.Core.Models; // Utilise le modèle SharePointSiteInfo global

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
        /// Commande asynchrone pour charger les sites SharePoint depuis Microsoft Graph
        /// </summary>
        [RelayCommand]

        private async Task LoadSitesAsync()
        {
            if (IsLoading) return;

            try
            {
                IsLoading = true;
                Sites.Clear();

                // Récupération du client Graph authentifié
                var graphClient = await _authService.GetAuthenticatedClientAsync();

                // Requête Microsoft Graph pour l'utilisateur connecté
                var sitesResult = await graphClient.Sites.GetAsync();

                if (sitesResult?.Value != null)
                {
                    foreach (var site in sitesResult.Value)
                    {
                        // CORRECTION SYNTAXE ICI : Bien vérifier les { } et ( )
                        Sites.Add(new SharePointSiteInfo
                        {
                            Id = site.Id ?? string.Empty,
                            DisplayName = site.DisplayName ?? site.Name ?? "Site sans nom"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Erreur lors du chargement des sites SharePoint : {ex.Message}",
                                "Erreur de chargement",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }


        /// <summary>
        /// Commande asynchrone pour valider la sélection du site
        /// </summary>
        [RelayCommand]
        private async Task SelectSiteAsync() // Changé en asynchrone
        {
            if (SelectedSite != null)
            {
                try
                {
                    // Enregistre le site sélectionné de manière asynchrone
                    await _selectionService.SetActiveSiteAsync(SelectedSite.Id);

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
                }
            }
        }
    }
}
