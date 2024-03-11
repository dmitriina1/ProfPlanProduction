using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using MaterialDesignThemes;

namespace ProfPlan.Views
{
    internal class TabStripPlacementConverter : IValueConverter
    {
        public ControlTemplate TopBottomTemplate { get; set; }
        public ControlTemplate LeftTemplate { get; set; }
        public ControlTemplate RightTemplate { get; set; }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Dock dockValue)
            {
                if (dockValue == Dock.Top || dockValue == Dock.Bottom)
                {
                    return TopBottomTemplate;
                }
                else if(dockValue == Dock.Left)
                {
                    return LeftTemplate;
                }
                else
                {
                    return RightTemplate;
                }
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
