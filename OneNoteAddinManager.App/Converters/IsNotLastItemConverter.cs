using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace OneNoteAddinManager.App.Converters
{
    /// <summary>
    /// Converter that returns Collapsed if the item is the last in the collection, Visible otherwise.
    /// Used to hide separators after the last item in a list.
    /// </summary>
    public class IsNotLastItemConverter : IMultiValueConverter
    {
        public static readonly IsNotLastItemConverter Instance = new();

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length != 2)
                return Visibility.Visible;

            var item = values[0];
            var collection = values[1] as IEnumerable;

            if (item == null || collection == null)
                return Visibility.Visible;

            var items = collection.Cast<object>().ToList();
            var isLast = items.LastOrDefault()?.Equals(item) == true;

            return isLast ? Visibility.Collapsed : Visibility.Visible;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}