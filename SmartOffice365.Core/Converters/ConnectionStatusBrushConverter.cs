using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Convertit l'état de connexion (bool) en couleur (Brush) : Vert si connecté, Rouge si déconnecté
    /// </summary>
    public class ConnectionStatusBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isConnected && isConnected)
            {
                return new SolidColorBrush(Color.FromRgb(0x3F, 0xB9, 0x50)); // Vert (#3FB950)
            }
            return new SolidColorBrush(Color.FromRgb(0xF8, 0x51, 0x49)); // Rouge (#F85149)
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
