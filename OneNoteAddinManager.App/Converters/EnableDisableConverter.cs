using System;
using System.Globalization;
using System.Windows.Data;

namespace OneNoteAddinManager.App.Converters
{
    public class EnableDisableConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isEnabled)
            {
                return isEnabled ? "Disable" : "Enable";
            }
            return "Enable";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
