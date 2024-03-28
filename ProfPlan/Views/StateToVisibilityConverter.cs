using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows;

namespace ProfPlan.Views
{
    public class StateToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Проверяем, является ли значение double и больше ли оно 0 (так как ProgressBar Minimum="0")
            if (value is double state && state > 0)
            {
                // Возвращаем Visible, если значение больше 0
                return Visibility.Visible;
            }
            else
            {
                // Возвращаем Collapsed, если значение не изменялось в течение 10 секунд или равно 0
                return Visibility.Collapsed;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
