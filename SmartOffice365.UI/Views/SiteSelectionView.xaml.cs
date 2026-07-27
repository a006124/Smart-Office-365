using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using SmartOffice365.UI.ViewModels;

namespace SmartOffice365.UI.Views
{
    /// <summary>
    /// Logique d'interaction pour SiteSelectionView.xaml
    /// </summary>
    public partial class SiteSelectionView : UserControl
    {
        public SiteSelectionView()
        {
            InitializeComponent();
            Loaded += OnLoaded;
        }

        public SiteSelectionViewModel? ViewModel => DataContext as SiteSelectionViewModel;

        public void SetViewModel(SiteSelectionViewModel viewModel)
        {
            DataContext = viewModel;
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            // Chargement automatique des sites lorsque la vue s'affiche
            if (ViewModel != null && !ViewModel.IsLoading && ViewModel.Sites.Count == 0)
            {
                await ViewModel.LoadSitesCommand.ExecuteAsync(null);
            }
        }

        // Gestion du double-clic sur un site pour le sélectionner
        private void ListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ViewModel?.SelectedSite != null)
            {
                ViewModel.SelectSiteCommand.Execute(null);
            }
        }

        // Gestion de la touche Entrée sur la liste pour sélectionner
        private void ListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && ViewModel?.SelectedSite != null)
            {
                ViewModel.SelectSiteCommand.Execute(null);
                e.Handled = true;
            }
        }
    }
}