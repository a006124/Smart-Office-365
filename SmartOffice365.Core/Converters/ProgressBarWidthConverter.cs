using System;
using System.Globalization;
using System.Windows.Data;

namespace SmartOffice365.Core.Converters
{
    /// <summary>
    /// Convertit un pourcentage en largeur pour une ProgressBar (0-100 -> 0-1)
    /// </summary>
    public class ProgressBarWidthConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int intValue)
                return Math.Max(0, Math.Min(1, intValue / 100.0));
            
            if (value is double doubleValue)
                return Math.Max(0, Math.Min(1, doubleValue / 100.0));
                
            return 0.0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}