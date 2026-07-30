using System;
using System.Globalization;
using System.Windows.Data;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Inverse une valeur booléenne. True devient False, False devient True
    /// </summary>
    public class InverseBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
                return !boolValue;
            return true;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
                return !boolValue;
            return false;
        }
    }
}