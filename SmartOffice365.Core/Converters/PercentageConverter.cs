using System;
using System.Globalization;
using System.Windows.Data;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Convertit un nombre en pourcentage (ex: 0.75 -> "75%")
    /// </summary>
    public class PercentageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double doubleValue)
                return $"{doubleValue:F0}%";
            
            if (value is int intValue)
                return $"{intValue}%";
            
            if (value is decimal decimalValue)
                return $"{decimalValue:F0}%";
                
            return "0%";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}