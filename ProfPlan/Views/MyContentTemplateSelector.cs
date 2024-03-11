using ProfPlan.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ProfPlan.Views
{
    internal class MyContentTemplateSelector : DataTemplateSelector
    {
        public DataTemplate FirstTemplate { get; set; }
        public DataTemplate SecondTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (!(item is TableCollection tableCollection)) return null;

            if (tableCollection.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1)
            {
                return FirstTemplate;
            }
            else
            {
                return SecondTemplate;
            }
        }
    }
}
