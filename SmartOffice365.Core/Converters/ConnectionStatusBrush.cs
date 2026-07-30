using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Convertit un statut de connexion booléen en couleur (Vert = connecté, Rouge = déconnecté)
    /// </summary>
    public class ConnectionStatusBrush : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isConnected)
            {
                return isConnected 
                    ? new SolidColorBrush(Colors.LightGreen) 
                    : new SolidColorBrush(Colors.Red);
            }
            return new SolidColorBrush(Colors.Gray);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}