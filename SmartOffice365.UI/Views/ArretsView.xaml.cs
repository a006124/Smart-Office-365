using System.Windows.Controls;
using SmartOffice365.UI.ViewModels;

namespace SmartOffice365.UI.Views
{
    public partial class ArretsView : System.Windows.Controls.UserControl
    {
        public ArretsView()
        {
            InitializeComponent();
            // Optionnel si déjà géré par le DataTemplate de MainWindow.xaml :
            // this.DataContext = new ArretsViewModel(); 
        }
    }
}
