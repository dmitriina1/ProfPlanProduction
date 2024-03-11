using Microsoft.Win32;
using ProfPlan.Commands;
using ProfPlan.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows;
using ExcelDataReader;
using System.ComponentModel;
using System.Windows.Controls;
using ProfPlan.ViewModels.Base;

namespace ProfPlan.ViewModels
{
    internal class MainViewModel : ViewModel
    {
        private string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Расчет нагрузки {DateTime.Today:dd-MM-yyyy}");
        private int Number = 1;
        private RelayCommand _loadDataCommand;
        private DataTableCollection tableCollection;
        public ICommand LoadDataCommand
        {
            get { return _loadDataCommand ?? (_loadDataCommand = new RelayCommand(LoadData)); }
        }
        private void LoadData(object parameter)
        {
            //try
            {
                string tabname = "";
                var openFileDialog = new OpenFileDialog() { Filter = "Excel Files|*.xls;*.xlsx" };
                if (openFileDialog.ShowDialog() == true)
                {
                    directoryPath = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = false }
                            });
                            tableCollection = result.Tables;
                            TablesCollections.Clear();

                            foreach (DataTable table in tableCollection)
                            {
                                Number = 1;
                                tabname = table.TableName;
                                ObservableCollection<ExcelData> list = new ObservableCollection<ExcelData>();
                                int rowIndex = -1;
                                bool haveTeacher = false;

                                //Определение индекса строки с заголовком "Дисциплина"
                                for (int i = 0; i < table.Rows.Count; i++)
                                {
                                    for (int j = 0; j < table.Columns.Count - 1; j++)
                                    {
                                        if (table.Rows[i][j].ToString().Trim() == "Дисциплина")
                                        {
                                            rowIndex = i;
                                            break;
                                        }
                                    }
                                }
                                bool exitOuterLoop = false;
                                int endstring = -1;
                                for (int i = 0; i < table.Rows.Count; i++)
                                {
                                    for (int j = 0; j < table.Columns.Count - 1; j++)
                                    {
                                        if (table.Rows[i][j].ToString().Trim() == "Дисциплина")
                                        {
                                            rowIndex = i;

                                            exitOuterLoop = true;
                                            break;
                                        }
                                    }

                                    if (exitOuterLoop)
                                    {
                                        // Выход из внешнего цикла
                                        break;
                                    }
                                }
                                if (rowIndex != -1)
                                    for (int i = rowIndex; i < table.Rows.Count; i++)
                                    {
                                        if (table.Rows[i][0].ToString() == "")
                                        {
                                            endstring = i;
                                            break;
                                        }
                                    }

                                // Проверка наличия столбца "Преподаватель"
                                for (int j = 0; j < table.Columns.Count - 1; j++)
                                {
                                    if (rowIndex != -1 && table.Rows[rowIndex][j].ToString().Trim() == "Преподаватель")
                                    {
                                        haveTeacher = true;
                                        break;
                                    }
                                }
                                var teachers = new ObservableCollection<string>();

                                // Заполнение коллекции данных
                                if (table.TableName.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1 &&
                                    table.TableName.IndexOf("доп", StringComparison.OrdinalIgnoreCase) == -1)
                                {

                                    if (endstring == -1) { endstring = table.Rows.Count; }
                                    for (int i = rowIndex + 1; i < endstring; i++)
                                    {
                                        try
                                        {
                                            if (haveTeacher && !string.IsNullOrWhiteSpace(table.Rows[i][0].ToString()))
                                            {
                                                list.Add(new ExcelModel(
                                                                       Number,
                                                                       table.Rows[i][1].ToString(),
                                                                       table.Rows[i][2].ToString(),
                                                                       table.Rows[i][3].ToString(),
                                                                       table.Rows[i][4].ToString(),
                                                                       table.Rows[i][5].ToString(),
                                                                       table.Rows[i][6].ToNullable<int>(),
                                                                       table.Rows[i][7].ToString(),
                                                                       table.Rows[i][8].ToString(),
                                                                       table.Rows[i][9].ToNullable<int>(),
                                                                       table.Rows[i][10].ToNullable<int>(),
                                                                       table.Rows[i][11].ToNullable<int>(),
                                                                       table.Rows[i][12].ToString(),
                                                                       table.Rows[i][13].ToNullable<int>(),
                                                                       table.Rows[i][14].ToNullable<double>(),
                                                                       table.Rows[i][15].ToNullable<double>(),
                                                                       table.Rows[i][16].ToNullable<double>(),
                                                                       table.Rows[i][17].ToNullable<double>(),
                                                                       table.Rows[i][18].ToNullable<double>(),
                                                                       table.Rows[i][19].ToNullable<double>(),
                                                                       table.Rows[i][20].ToNullable<double>(),
                                                                       table.Rows[i][21].ToNullable<double>(),
                                                                       table.Rows[i][22].ToNullable<double>(),
                                                                       table.Rows[i][23].ToNullable<double>(),
                                                                       table.Rows[i][24].ToNullable<double>(),
                                                                       table.Rows[i][25].ToNullable<double>(),
                                                                       table.Rows[i][26].ToNullable<double>(),
                                                                       table.Rows[i][27].ToNullable<double>(),
                                                                       table.Rows[i][28].ToNullable<double>()));
                                                Number++;
                                            }
                                            else if (!haveTeacher)
                                            {
                                                list.Add(new ExcelModel(
                                                                       Number,
                                                                       "",
                                                                       table.Rows[i][1].ToString(),
                                                                       table.Rows[i][2].ToString(),
                                                                       table.Rows[i][3].ToString(),
                                                                       table.Rows[i][4].ToString(),
                                                                       table.Rows[i][5].ToNullable<int>(),
                                                                       table.Rows[i][6].ToString(),
                                                                       table.Rows[i][7].ToString(),
                                                                       table.Rows[i][8].ToNullable<int>(),
                                                                       table.Rows[i][9].ToNullable<int>(),
                                                                       table.Rows[i][10].ToNullable<int>(),
                                                                       table.Rows[i][11].ToString(),
                                                                       table.Rows[i][12].ToNullable<double>(),
                                                                       table.Rows[i][13].ToNullable<double>(),
                                                                       table.Rows[i][14].ToNullable<double>(),
                                                                       table.Rows[i][15].ToNullable<double>(),
                                                                       table.Rows[i][16].ToNullable<double>(),
                                                                       table.Rows[i][17].ToNullable<double>(),
                                                                       table.Rows[i][18].ToNullable<double>(),
                                                                       table.Rows[i][19].ToNullable<double>(),
                                                                       table.Rows[i][20].ToNullable<double>(),
                                                                       table.Rows[i][21].ToNullable<double>(),
                                                                       table.Rows[i][22].ToNullable<double>(),
                                                                       table.Rows[i][23].ToNullable<double>(),
                                                                       table.Rows[i][24].ToNullable<double>(),
                                                                       table.Rows[i][25].ToNullable<double>(),
                                                                       table.Rows[i][26].ToNullable<double>(),
                                                                       table.Rows[i][27].ToNullable<double>()));
                                                Number++;
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Error adding data: {ex.Message}");
                                        }
                                    }
                                }
                                else if (table.TableName.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
                                {
                                    bool hasBetPer = false;
                                    for (int i = 1; i < table.Columns.Count; i++)
                                    {
                                        if (table.Rows[0][i].ToString().IndexOf("%", StringComparison.OrdinalIgnoreCase) != -1)
                                        {
                                            hasBetPer = true;
                                            break;
                                        }
                                    }
                                    if (hasBetPer != true)
                                        for (int i = 1; i < table.Rows.Count; i++)
                                        {
                                            if (!string.IsNullOrEmpty(table.Rows[i][0].ToString()))
                                                list.Add(new ExcelTotal(
                                                    table.Rows[i][0].ToString(),
                                                    table.Rows[i][1].ToNullable<int>(),
                                                    null,
                                                    table.Rows[i][2].ToNullable<double>(),
                                                    table.Rows[i][3].ToNullable<double>(),
                                                    table.Rows[i][4].ToNullable<double>(),
                                                    Math.Round(Convert.ToDouble(table.Rows[i][5].ToNullable<double>()), 2)
                                                    ));
                                        }
                                    else
                                    {
                                        for (int i = 1; i < table.Rows.Count; i++)
                                        {
                                            if (!string.IsNullOrEmpty(table.Rows[i][0].ToString()))
                                                list.Add(new ExcelTotal(
                                                    table.Rows[i][0].ToString(),
                                                     table.Rows[i][1].ToNullable<int>(),
                                                    table.Rows[i][2].ToNullable<double>(),
                                                    table.Rows[i][3].ToNullable<double>(),
                                                    table.Rows[i][4].ToNullable<double>(),
                                                    table.Rows[i][5].ToNullable<double>(),
                                                    table.Rows[i][6].ToNullable<double>()
                                                    ));
                                        }
                                    }

                                }

                                // Добавление коллекции в TablesCollection
                                TablesCollections.Add(new TableCollection(tabname, list));
                            }
                        }
                    }

                    // Обновление свойства привязок данных в XAML
                    OnPropertyChanged(nameof(TablesCollections));
                    UpdateListBoxItemsSource();
                }
            }
        }

        // Все для взаимодействия ComboBox и ListBox

        private int _selectedComboBoxIndex;

        public int SelectedComboBoxIndex
        {
            get { return _selectedComboBoxIndex; }
            set
            {
                if (_selectedComboBoxIndex != value)
                {
                    _selectedComboBoxIndex = value;
                    OnPropertyChanged(nameof(SelectedComboBoxIndex));

                    // Обновляем ItemsSource для ListBox в зависимости от выбранного элемента в ComboBox
                    UpdateListBoxItemsSource();
                }
            }
        }
        private ObservableCollection<TableCollection> _displayedTables;

        public ObservableCollection<TableCollection> DisplayedTables
        {
            get { return _displayedTables; }
            set
            {
                if (_displayedTables != value)
                {
                    _displayedTables = value;
                    OnPropertyChanged(nameof(DisplayedTables));
                }
            }
        }
        private void UpdateListBoxItemsSource()
        {
            if (SelectedComboBoxIndex == 0)
            {
                DisplayedTables = TablesCollections.GetTablesCollectionWithP();
            }
            else if (SelectedComboBoxIndex == 1)
            {
                DisplayedTables = TablesCollections.GetTablesCollectionWithF();
            }
        }

        //Все для взаимодействия listbox и datagrid

        private TableCollection _selectedTable;

        public TableCollection SelectedTable
        {
            get { return _selectedTable; }
            set
            {
                if (_selectedTable != value)
                {
                    _selectedTable = value;
                    OnPropertyChanged(nameof(SelectedTable));
                    // Обновляем данные в DataGrid при выборе нового элемента в ListBox
                    


                }
            }
        }

        private Dock _tabStripPlacement = Dock.Top;

        public Dock TabStripPlacement
        {
            get { return _tabStripPlacement; }
            set
            {
                if (_tabStripPlacement != value)
                {
                    _tabStripPlacement = value;
                    OnPropertyChanged(nameof(TabStripPlacement));
                }
            }
        }
        // Выбор содержимого dataGrid


    }   
}
