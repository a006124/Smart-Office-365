using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace SmartOffice365.UI.Converters
{
    /// <summary>
    /// Convertit une priorité en couleur
    /// </summary>
    public class PriorityToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var priority = value as string ?? string.Empty;
            
            switch (priority.ToLower())
            {
                case "critique":
                    return new SolidColorBrush(Colors.Red);
                case "haute":
                    return new SolidColorBrush(Colors.Orange);
                case "normale":
                    return new SolidColorBrush(Colors.DodgerBlue);
                case "basse":
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