using ProfPlan.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlan.Models
{
    internal static class TablesCollections 
    {
        private static ObservableCollection<TableCollection> TablesCollection = new ObservableCollection<TableCollection>();
        
        public static void Clear()
        {
            TablesCollection.Clear();
        }
        public static void Add(TableCollection tabCol)
        {
            TablesCollection.Add(tabCol);
        }

        public static ObservableCollection<TableCollection> GetTablesCollection()
        {
            return new ObservableCollection<TableCollection>(TablesCollection);
        }

        public static ObservableCollection<TableCollection> GetTablesCollectionWithP()
        {
            return new ObservableCollection<TableCollection>(
                TablesCollection.Where(tc => tc.Tablename.StartsWith("П_")).ToList());
        }

        public static ObservableCollection<TableCollection> GetTablesCollectionWithF()
        {
            return new ObservableCollection<TableCollection>(
                TablesCollection.Where(tc => tc.Tablename.StartsWith("Ф_")).ToList());
        }

        public static bool GetTableByName(string tableName, int selectedIndex)
        {
            if(selectedIndex == 0)
            {
                foreach (TableCollection table in TablesCollections.GetTablesCollectionWithP())
                {
                    if (table.Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return true;
                    }
                }
            }
            else if(selectedIndex == 1)
            {
                foreach (TableCollection table in TablesCollections.GetTablesCollectionWithF())
                {
                    if (table.Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return true;
                    }
                }
            }
            return false;

        }

    }
}
