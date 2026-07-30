using System;
using System.Globalization;
using System.Windows.Data;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Convertit un booléen en texte de statut (True = "Actif", False = "Inactif")
    /// </summary>
    public class BoolToStatusConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                var trueText = "Actif";
                var falseText = "Inactif";
                
                // Permet de personnaliser les textes via le paramètre
                if (parameter is string param && param.Contains('|'))
                {
                    var parts = param.Split('|');
                    if (parts.Length >= 2)
                    {
                        trueText = parts[0];
                        falseText = parts[1];
                    }
                }
                
                return boolValue ? trueText : falseText;
            }
            return "Inconnu";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}