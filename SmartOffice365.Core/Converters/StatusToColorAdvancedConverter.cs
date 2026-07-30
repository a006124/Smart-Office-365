using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Version avancée : Convertit un statut en couleur spécifique
    /// </summary>
    public class StatusToColorAdvancedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var status = value as string ?? string.Empty;
            
            // Couleurs pour différents statuts
            switch (status.ToLower())
            {
                case var s when s.Contains("critique") || s.Contains("bloqué"):
                    return new SolidColorBrush(Colors.Red);
                    
                case var s when s.Contains("retard"):
                    return new SolidColorBrush(Colors.Orange);
                    
                case var s when s.Contains("en cours"):
                    return new SolidColorBrush(Colors.DodgerBlue);
                    
                case var s when s.Contains("terminé"):
                    return new SolidColorBrush(Colors.LightGreen);
                    
                case var s when s.Contains("planifié") || s.Contains("à faire"):
                    return new SolidColorBrush(Colors.Gray);
                    
                default:
                    return new SolidColorBrush(Colors.LightGray);
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}