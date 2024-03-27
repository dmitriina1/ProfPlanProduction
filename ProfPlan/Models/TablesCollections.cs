using ProfPlan.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProfPlan.Models
{
    internal static class TablesCollections 
    {
        private static ObservableCollection<TableCollection> TablesCollection = new ObservableCollection<TableCollection>();
        public static void AddByIndex(int index, ExcelData tab)
        {
            TablesCollection[index].ExcelData.Add(tab);
        }
        public static int Count()
        {
            return TablesCollection.Count();
        }
        public static void Clear()
        {
            TablesCollection.Clear();
        }
        public static void Add(TableCollection tabCol)
        {
            int foundIndex = GetTableIndexByName(tabCol.Tablename);
            if (foundIndex != -1)
            {
                TablesCollection[foundIndex] = tabCol;
            }
            else
            {
                TablesCollection.Add(tabCol);
            }
        }
        public static void AddInOldTabCol(TableCollection tabCol)
        {
            int foundIndex = GetTableIndexByName(tabCol.Tablename);

            if (foundIndex != -1)
            {
                for(int i=0;i<tabCol.ExcelData.Count;i++)
                {
                    TablesCollection[foundIndex].ExcelData.Add(tabCol.ExcelData[i]);
                }
            }
            else
            {
                TablesCollection.Add(tabCol);
            }
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
            if (selectedIndex == 0)
            {
                foreach (TableCollection table in TablesCollections.GetTablesCollectionWithP())
                {
                    if (table.Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return true;
                    }
                }
            }
            else if (selectedIndex == 1)
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
        public static int GetTableIndexByName(string tableName, int selectedIndex = -2)
        {

            if (selectedIndex == -1)
            {
                return -1; // Некорректное значение selectedIndex
            }
            
                for (int i = 0; i < TablesCollections.TablesCollection.Count; i++)
                {
                    if (string.Equals(TablesCollections.TablesCollection[i].Tablename, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return i; // Возвращаем индекс, если таблица найдена
                    }
                }
            
            return -1;

        }

        public static int GetTableIndexForGenerate(string tableName, int selectedIndex)
        {

            if (selectedIndex == -1)
            {
                return -1; // Некорректное значение selectedIndex
            }
            if (selectedIndex == 0)
            {
                for (int i = 0; i < TablesCollections.GetTablesCollectionWithP().Count; i++)
                {
                    if (TablesCollections.GetTablesCollectionWithP()[i].Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return i; // Возвращаем индекс, если таблица найдена
                    }
                }
            }
            else
            {
                for (int i = 0; i < TablesCollections.GetTablesCollectionWithF().Count; i++)
                {
                    if (TablesCollections.GetTablesCollectionWithF()[i].Tablename.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) != -1)
                    {
                        return i; // Возвращаем индекс, если таблица найдена
                    }
                }
            }

            return -1;

        }

        public static void RemoveTableAtIndex(int index)
        {
            if (index >= 0 && index < TablesCollection.Count)
            {
                TablesCollection[index].ExcelData.Clear();
            }
        }

        public static void Insert(int index, TableCollection tabCol)
        {
            TablesCollection.Insert(index, tabCol);
        }


        public static void SortTablesCollection()
        {
            var sortedCollectionP = TablesCollections.GetTablesCollectionWithP().OrderBy(tc =>
            {
                if (tc.Tablename.Contains("ПИиИС")) return 0;
                if (tc.Tablename.Contains("Итого")) return 1;
                if (tc.Tablename.Contains("Доп")) return 2;
                return 3;
            }).ThenBy(tc => tc.Tablename).ToList();

            var sortedCollectionF = TablesCollections.GetTablesCollectionWithF().OrderBy(tc =>
            {
                if (tc.Tablename.Contains("ПИиИС")) return 0;
                if (tc.Tablename.Contains("Итого")) return 1;
                if (tc.Tablename.Contains("Доп")) return 2;
                return 3;
            }).ThenBy(tc => tc.Tablename).ToList();

            TablesCollection.Clear();
            foreach (var table in sortedCollectionP)
            {
                TablesCollection.Add(table);
            }
            foreach (var table in sortedCollectionF)
            {
                TablesCollection.Add(table);
            }
        }
    }
}
