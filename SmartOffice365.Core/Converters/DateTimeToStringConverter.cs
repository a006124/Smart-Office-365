using System;
using System.Globalization;
using System.Windows.Data;

namespace SmartOffice365.UI.Converters
{
    /// <summary>
    /// Convertit un DateTime en chaîne formatée
    /// </summary>
    public class DateTimeToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is DateTime dateTime)
            {
                var format = parameter as string ?? "dd/MM/yyyy HH:mm";
                return dateTime.ToString(format, CultureInfo.CurrentCulture);
            }
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}