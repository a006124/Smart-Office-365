using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Convertit une valeur null en Visibility inversé. Null = Visible, Non-null = Collapsed
    /// </summary>
    public class InverseNullToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value == null ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}