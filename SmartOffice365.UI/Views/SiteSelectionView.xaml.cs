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
    public partial class SiteSelectionView : System.Windows.Controls.UserControl
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
            // Le code est maintenant réactivé et fonctionnel !
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
        private void ListBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter && ViewModel?.SelectedSite != null)
            {
                ViewModel.SelectSiteCommand.Execute(null);
                e.Handled = true;
            }
        }
    }
}