using System;
using System.Globalization;
using System.Windows.Data;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Vérifie si une valeur est null. Retourne True si la valeur n'est pas null
    /// </summary>
    public class HasValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}