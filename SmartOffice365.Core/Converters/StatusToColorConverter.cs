using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace SmartOffice365.UI.Converters
{
    /// <summary>
    /// Convertit un message de statut en couleur en fonction de son contenu
    /// </summary>
    public class StatusToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var status = value as string ?? string.Empty;
            
            if (status.Contains("Erreur") || status.Contains("⚠️") || status.Contains("échec"))
                return new SolidColorBrush(Colors.OrangeRed);
            
            if (status.Contains("succès") || status.Contains("✅") || status.Contains("terminé"))
                return new SolidColorBrush(Colors.LightGreen);
            
            if (status.Contains("Chargement") || status.Contains("...") || status.Contains("Vérification"))
                return new SolidColorBrush(Colors.LightGray);
            
            if (status.Contains("Prêt") || status.Contains("Connecté"))
                return new SolidColorBrush(Colors.LightGreen);
                
            return new SolidColorBrush(Colors.Gray);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}