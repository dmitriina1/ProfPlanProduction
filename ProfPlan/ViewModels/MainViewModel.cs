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
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ProfPlan.Views;
using System.Collections.Specialized;

namespace ProfPlan.ViewModels
{
    internal class MainViewModel : ViewModel
    {
        public MainViewModel()
        {
            ExcelModel.UpdateSharedTeachers();
        }

        private string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Расчет нагрузки {DateTime.Today:dd-MM-yyyy}");
        private string filePath = "", tempFilePath = "";
        private int Number = 1;
        private DataTableCollection tableCollection;

        /// <summary>
        /// Вкладка Файл 
        /// </summary>

        // Открытие Excel файла
        #region Open Data from Excel
        private RelayCommand _loadDataCommand;

        public ICommand LoadDataCommand
        {
            get { return _loadDataCommand ?? (_loadDataCommand = new RelayCommand(LoadData)); }
        }

        private void LoadData(object parameter)
        {
            filePath = GetExcelFilePath();
            if (!string.IsNullOrEmpty(filePath))
            {
                tableCollection = ReadExcelData(filePath).Tables;
                TablesCollections.Clear();
                foreach (DataTable table in tableCollection)
                {
                    ProcessDataTable(table);
                }
                OnPropertyChanged(nameof(TablesCollections));
                UpdateListBoxItemsSource();
            }

        }

        private string GetExcelFilePath()
        {
            var openFileDialog = new OpenFileDialog() { Filter = "Excel Files|*.xls;*.xlsx" };

            return openFileDialog.ShowDialog() == true ? openFileDialog.FileName : null;
        }

        private DataSet ReadExcelData(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = false }
                    });
                }
            }
        }

        private void ProcessDataTable(DataTable table)
        {
            Number = 1;
            string tabname = table.TableName;
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
                        //MessageBox.Show($"Error adding data: {ex.Message}");
                    }
                }
            }
            else if (table.TableName.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
            {
                ProcessTotalTable(table, list);
            }
            for(int i=0;i<list.Count; i++)
            {
                list[i].PropertyChanged +=SelectedItemPropertyChanged;
            }
            TablesCollections.Add(new TableCollection(tabname, list));

        }

        private void ProcessTotalTable(DataTable table, ObservableCollection<ExcelData> list)
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

        #endregion

        //Добавление данных из Excel-файла
        #region Add Data from Excel
        private RelayCommand _addDataCommand;

        public ICommand AddDataCommand
        {
            get { return _addDataCommand ?? (_addDataCommand = new RelayCommand(AddData)); }
        }
        private void AddData(object parameter)
        {
            tempFilePath = GetExcelFilePath();
            if (!string.IsNullOrEmpty(tempFilePath))
            {
                tableCollection = ReadExcelData(tempFilePath).Tables;
                
                if (tableCollection.Count == 1)
                {
                    if (_selectedComboBoxIndex == 0)
                        tableCollection[0].TableName = "П_ПИиИС";
                    else if(_selectedComboBoxIndex == 1)
                        tableCollection[0].TableName = "Ф_ПИиИС";
                    foreach (DataTable table in tableCollection)
                    {
                        DataTableInsert(table);
                    }
                    OnPropertyChanged(nameof(TablesCollections));
                    UpdateListBoxItemsSource();
                }
                else
                {
                    MessageBox.Show("Ошибка! Можно добавить лишь 1 таблицу!");
                }
                
            }

        }

        private void DataTableInsert(DataTable table)
        {
            int ind = -1;
            if (_selectedComboBoxIndex == 0)
                ind = TablesCollections.GetTableIndexByName("П_ПИиИС", _selectedComboBoxIndex);
            else if (_selectedComboBoxIndex == 1)
                ind = TablesCollections.GetTableIndexByName("П_ПИиИС", _selectedComboBoxIndex);
            if(ind == -1)
            {
                Number = 1;
            }
            else
            {
                Number = TablesCollections.GetTablesCollection()[ind].ExcelData.Count() + 1;
            }
            string tabname = table.TableName;
            ObservableCollection<ExcelData> list = new ObservableCollection<ExcelData>();
            int rowIndex = -1;
            bool haveTeacher = false;
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
            for (int j = 0; j < table.Columns.Count - 1; j++)
            {
                if (rowIndex != -1 && table.Rows[rowIndex][j].ToString().Trim() == "Преподаватель")
                {
                    haveTeacher = true;
                    break;
                }
            }

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
                        //MessageBox.Show($"Error adding data: {ex.Message}");
                    }
                }
            
            for (int i = 0; i<list.Count; i++)
            {
                list[i].PropertyChanged +=SelectedItemPropertyChanged;
            }
            TablesCollections.AddInOldTabCol(new TableCollection(tabname, list));
        }

        #endregion

        // Обновление содержимого dataGrid и выбор отображаемых таблиц в tabControl
        #region Display data in a datagrid
        private int _selectedComboBoxIndex;
        private ObservableCollection<TableCollection> _displayedTables;
        private TableCollection _selectedTable;
        private ObservableCollection<ExcelData> _selectedItems = new ObservableCollection<ExcelData>();
        public ObservableCollection<ExcelData> SelectedItems
        { get { return _selectedItems; } }

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

        private void SelectedItemPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Teacher")
            {
                var changedItem = (ExcelModel)sender;
                var newTeacher = changedItem.Teacher;

                foreach (ExcelModel item in _selectedItems)
                {
                    if (item != changedItem)
                    {
                        item.Teacher = newTeacher;
                    }
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

        public TableCollection SelectedTable
        {
            get { return _selectedTable; }
            set
            {
                if (_selectedTable != value)
                {
                    _selectedTable = value;
                    OnPropertyChanged(nameof(SelectedTable));
                }
            }
        }


        #endregion

        // Выбор места для отображения tabItems
        #region Settings
        private Dock _tabStripPlacement = Dock.Left;
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
        private string _placementIcon = "ArrowLeft";
        public string PlacementIcon
        {
            get { return _placementIcon; }
            set
            {
                if (_placementIcon != value)
                {
                    _placementIcon = value;
                    OnPropertyChanged(nameof(PlacementIcon));
                }
            }
        }

        private RelayCommand _selectTabItemsPlacement;
        public ICommand SelectTabItemsPlacementCommand
        {
            get { return _selectTabItemsPlacement ?? (_selectTabItemsPlacement = new RelayCommand(SelectTabItemsPlacement)); }
        }
        private void SelectTabItemsPlacement(object parameter)
        {
            switch (_tabStripPlacement)
            {
                case Dock.Top:
                    _tabStripPlacement = Dock.Right;
                    PlacementIcon = "ArrowRight";
                    break;
                case Dock.Right:
                    _tabStripPlacement = Dock.Bottom;
                    PlacementIcon = "ArrowDown";
                    break;
                case Dock.Bottom:
                    _tabStripPlacement = Dock.Left;
                    PlacementIcon = "ArrowLeft";
                    break;
                case Dock.Left:
                    _tabStripPlacement = Dock.Top;
                    PlacementIcon = "ArrowUp";
                    break;
                default:
                    _tabStripPlacement = Dock.Left;
                    PlacementIcon = "ArrowLeft";
                    break;
            }
            OnPropertyChanged(nameof(TabStripPlacement));
            OnPropertyChanged(nameof(PlacementIcon));
        }
        #endregion

        // Сохраение Расчета нагрузки
        #region Save and SaveAs
        private RelayCommand _saveDataToExcel;
        private RelayCommand _saveDataToExcelAs;

        public ICommand SaveDataCommand
        {
            get { return _saveDataToExcel ?? (_saveDataToExcel = new RelayCommand(SaveToExcel)); }
        }

        private void SaveToExcel(object parameter)
        {
            if(filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) == false || filePath == "")
                SaveToExcelAs();
            else
            SaveToExcels(TablesCollections.GetTablesCollection());
        }

        private void SaveToExcels(ObservableCollection<TableCollection> tablesCollection)
        {
            using (var workbook = new XLWorkbook())
            {
                foreach (var table in tablesCollection)
                {
                    var worksheet = CreateWorksheet(workbook, table);
                    PopulateWorksheet(worksheet, table);
                }
                int frow = 2;
                List<string> newPropertyNames = new List<string>
    {
        "№","Преподаватель", "Дисциплина","Семестр(четный или нечетный)","Группа","Институт","Число групп","Подгруппа","Форма обучения","Число студентов","Из них коммерч.","Недель","Форма отчетности","Лекции",  "Практики","Лабораторные","Консультации", "Зачеты", "Экзамены", "Курсовые работы", "Курсовые проекты",  "ГЭК+ПриемГЭК, прием ГАК",
        "Диплом","РГЗ_Реф, нормоконтроль","ПрактикаРабота, реценз диплом", "Прочее", "Всего","Бюджетные","Коммерческие"
    };
                List<string> newPropertyTotalNames = new List<string>
    {
        "ФИО","Ставка", "Ставка(%)","Всего","Осень","Весна","Разница"
    };
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1)
                    {
                        for (int i = 0; i < newPropertyNames.Count; i++)
                        {
                            worksheet.Cell(frow, i + 1).Value = newPropertyNames[i];
                            worksheet.Cell(frow, i + 1).Style.Alignment.SetTextRotation(90);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < newPropertyTotalNames.Count; i++)
                        {
                            worksheet.Cell(frow - 1, i + 1).Value = newPropertyTotalNames[i];
                        }
                    }
                }
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1)
                    {
                        worksheet.Rows().AdjustToContents();
                    }
                }


                SaveWorkbook(workbook);
            }
        }

        private IXLWorksheet CreateWorksheet(XLWorkbook workbook, TableCollection table)
        {
            var worksheet = workbook.Worksheets.Add(table.Tablename);

            if (table.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
            {
                CreateTotalHeaders(worksheet);
            }
            else
            {
                CreateModelHeaders(worksheet);
            }

            return worksheet;
        }

        private void CreateTotalHeaders(IXLWorksheet worksheet)
        {
            int columnNumber = 1;
            foreach (var propertyInfo in typeof(ExcelTotal).GetProperties())
            {
                worksheet.Cell(1, columnNumber).Value = propertyInfo.Name;
                columnNumber++;
            }
        }

        private void CreateModelHeaders(IXLWorksheet worksheet)
        {
            int columnNumber = 1;
            foreach (var propertyInfo in typeof(ExcelModel).GetProperties())
            {
                if (propertyInfo.Name != "Teachers")
                {
                    worksheet.Cell(2, columnNumber).Value = propertyInfo.Name;
                    columnNumber++;
                }
            }
        }

        private void PopulateWorksheet(IXLWorksheet worksheet, TableCollection table)
        {
            int rowNumber = table.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1 ? 2 : 3;
            int columnNumber = 1;

            foreach (var data in table.ExcelData)
            {
                foreach (var propertyName in GetPropertyNames(data))
                {
                    var value = data.GetType().GetProperty(propertyName)?.GetValue(data, null);
                    worksheet.Cell(rowNumber, columnNumber).Value = value != null ? value.ToString() : "";
                    columnNumber++;
                }

                rowNumber++;
                columnNumber = 1;
            }

        }

        private IEnumerable<string> GetPropertyNames(object data)
        {
            return data is ExcelModel model
                ? typeof(ExcelModel).GetProperties().Where(p => p.Name != "Teachers").Select(p => p.Name)
                : typeof(ExcelTotal).GetProperties().Select(p => p.Name);
        }

        private void SaveWorkbook(XLWorkbook workbook)
        {
            workbook.SaveAs(directoryPath);
        }

        public ICommand SaveDataAsCommand
        {
            get { return _saveDataToExcelAs ?? (_saveDataToExcelAs = new RelayCommand(SaveToExcelAs)); }
        }

        private void SaveToExcelAs(object parameter)
        {
            SaveToExcelAs();
        }

        private void SaveToExcelAs()
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.FileName = $"Расчет Нагрузки {DateTime.Today:dd-MM-yyyy}.xlsx";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                directoryPath = saveFileDialog.FileName;
            }
            else
            {
                return;
            }
            SaveToExcels(TablesCollections.GetTablesCollection());
        }

        #endregion

        //Выход из приложения
        #region Exit
        private RelayCommand _exitCommand;

        public ICommand ExitCommand
        {
            get { return _exitCommand ?? (_exitCommand = new RelayCommand(ExitFromApp)); }
        }
        private void ExitFromApp(object parameter)
        {
            Application.Current.Dispatcher.Invoke(() => Application.Current.Shutdown());
        }
        #endregion

        //Создать таблицу для плана или факта
        #region Create table
        private RelayCommand _createCommand;

        public ICommand CreateCommand
        {
            get { return _createCommand ?? (_createCommand = new RelayCommand(CreateBaseTableCollection)); }
        }
        private void CreateBaseTableCollection(object parameter)
        {
            CreateTableCollection();
        }
        private void CreateTableCollection()
        {
            if (SelectedComboBoxIndex == 0 && TablesCollections.GetTableByName("ПИиИС", SelectedComboBoxIndex) == false)
            {
                TablesCollections.Add(new TableCollection() { Tablename = "П_ПИиИС" });
            }
            else if (SelectedComboBoxIndex == 1 && TablesCollections.GetTableByName("ПИиИС", SelectedComboBoxIndex) == false)
            {
                TablesCollections.Add(new TableCollection() { Tablename = "Ф_ПИиИС" });
            }
            UpdateListBoxItemsSource();
        }
        #endregion



        /// <summary>
        /// Вкладка Таблица
        /// </summary>

        //Очистить таблицу
        #region Clear table
        private RelayCommand _clearTableCommand;

        public ICommand ClearTableCommand
        {
            get { return _clearTableCommand ?? (_clearTableCommand = new RelayCommand(ClearTable)); }
        }
        private void ClearTable(object parameter)
        {
            if(SelectedTable != null && SelectedComboBoxIndex != -1)
            {
                TablesCollections.RemoveTableAtIndex(TablesCollections.GetTableIndexByName(SelectedTable.Tablename, SelectedComboBoxIndex));
                UpdateListBoxItemsSource();
            }
        }
        #endregion

        //Перенести преподавателей
        #region Move Teachers from Plan to Fact
        private RelayCommand _moveTeachersCommand;

        public ICommand MoveTeachersCommand
        {
            get { return _moveTeachersCommand ?? (_moveTeachersCommand = new RelayCommand(MoveTeachers)); }
        }

        private void MoveTeachers(object parameter)
        {
            int ftableindex = TablesCollections.GetTableIndexByName("П_ПИиИС", SelectedComboBoxIndex);
            int stableindex = TablesCollections.GetTableIndexByName("Ф_ПИиИС", SelectedComboBoxIndex);
            if (ftableindex != -1 && stableindex != -1)
            {
                try
                {
                    if (TablesCollections.GetTablesCollection()[stableindex].ExcelData.Count != 0)
                    {
                        for (int i = 0; i < TablesCollections.GetTablesCollection()[stableindex].ExcelData.Count; i++)
                        {
                            if (TablesCollections.GetTablesCollection()[stableindex].ExcelData[i] is ExcelModel excelModel && excelModel.Teacher == "")
                            {
                                ExcelModel stableData = TablesCollections.GetTablesCollection()[stableindex].ExcelData[i] as ExcelModel;
                                ExcelModel ftableData = TablesCollections.GetTablesCollection()[ftableindex].ExcelData[i] as ExcelModel;

                                if (stableData != null && ftableData != null &&
                                    stableData.Term == ftableData.Term &&
                                    stableData.Group == ftableData.Group &&
                                    stableData.Institute == ftableData.Institute &&
                                    stableData.FormOfStudy == ftableData.FormOfStudy &&
                                    ftableData.Teacher != "")
                                {
                                    stableData.Teacher = ftableData.Teacher;
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Лист Факт пустой! Поэтому данные с листа План были скопированы");
                        CreateTableCollectionsForMove();
                        ftableindex = TablesCollections.GetTableIndexByName("П_ПИиИС", SelectedComboBoxIndex);
                        stableindex = TablesCollections.GetTableIndexByName("Ф_ПИиИС", SelectedComboBoxIndex);
                        for (int i = 0; i < TablesCollections.GetTablesCollection()[ftableindex].ExcelData.Count; i++)
                        {
                            ExcelModel ftableData = TablesCollections.GetTablesCollection()[ftableindex].ExcelData[i] as ExcelModel;
                            TablesCollections.AddByIndex(stableindex, ftableData);
                        }
                    }

                }
                catch
                {

                }
            }
            else
            {
                MessageBox.Show("Лист Факт пустой! Поэтому данные с листа План были скопированы");
                CreateTableCollectionsForMove();
                ftableindex = TablesCollections.GetTableIndexByName("П_ПИиИС", SelectedComboBoxIndex);
                stableindex = TablesCollections.GetTableIndexByName("Ф_ПИиИС", SelectedComboBoxIndex);
                for (int i = 0; i < TablesCollections.GetTablesCollection()[ftableindex].ExcelData.Count; i++)
                {
                    ExcelModel ftableData = TablesCollections.GetTablesCollection()[ftableindex].ExcelData[i] as ExcelModel;
                    TablesCollections.AddByIndex(stableindex, ftableData);
                }
            }
            SelectedTable = null;
            UpdateListBoxItemsSource();
        }

        private void CreateTableCollectionsForMove()
        {
            if (TablesCollections.GetTableByName("П_ПИиИС", 0) == false)
            {
                TablesCollections.Add(new TableCollection() { Tablename = "П_ПИиИС" });
            }
            if ( TablesCollections.GetTableByName("Ф_ПИиИС", 1) == false)
            {
                TablesCollections.Add(new TableCollection() { Tablename = "Ф_ПИиИС" });
            }
            UpdateListBoxItemsSource();
        }
        #endregion

        //Генерация листов преподавателей
        #region Generate Teachers lists
        private RelayCommand _generateTeachersLists;

        public ICommand GenerateTeachersLists
        {
            get { return _generateTeachersLists ?? (_generateTeachersLists = new RelayCommand(GenerateTeacher)); }
        }

        private void GenerateTeacher(object parameter)
        {
            if(SelectedComboBoxIndex != -1 && TablesCollections.GetTableIndexForGenerate("ПИиИС", SelectedComboBoxIndex) != -1)
            {
                string prefix;
                if(SelectedComboBoxIndex == 0)
                {
                    prefix = "П_";
                }
                else
                {
                    prefix = "Ф_";
                }
                int mainList = TablesCollections.GetTableIndexByName(prefix + "ПИиИС", SelectedComboBoxIndex);
                var uniqueTeachers = TablesCollections.GetTablesCollection()[mainList].ExcelData
               .Where(data => data is ExcelModel) // Фильтрация по типу ExcelModel
               .Select(data => ((ExcelModel)data).Teacher) // Приведение к ExcelModel и выбор Teacher
               .Distinct()
               .ToList();
                ObservableCollection<ExcelData> totallist = new ObservableCollection<ExcelData>();
                foreach (var teacher in uniqueTeachers)
                {
                    var teacherTableCollection = new TableCollection() { };

                    if (teacher.ToString() != "")
                        teacherTableCollection = new TableCollection(prefix+teacher.ToString().Split(' ')[0]);
                    else
                        teacherTableCollection = new TableCollection(prefix+"Незаполненные");
                    var teacherRows = TablesCollections.GetTablesCollection()[mainList].ExcelData
                    .Where(data => data is ExcelModel && ((ExcelModel)data).Teacher == teacher)
                    .ToList();
                    foreach (ExcelModel techrow in teacherRows)
                    {
                        techrow.PropertyChanged += teacherTableCollection.ExcelModel_PropertyChanged;
                        teacherTableCollection.ExcelData.Add(techrow);

                    }
                    teacherTableCollection.SubscribeToExcelDataChanges();
                    TablesCollections.Add(teacherTableCollection);
                    //Реализация листа Итого:

                    double? bet = null;
                    string lname, fname, mname;
                    if (teacherTableCollection.Tablename != prefix + "Незаполненные")
                    {
                        foreach(Teacher teach in TeachersManager.GetTeachers())
                        {
                            lname=teach.LastName;
                            fname=teach.FirstName;
                            mname=teach.MiddleName;
                            if($"{lname} {fname[0]}.{mname[0]}." == teacher)
                            {
                                bet = teach.Workload;
                            }

                        }
                        totallist.Add(new ExcelTotal(
                        teacher.IndexOf(' ') != -1 ? teacher.Substring(0, teacher.IndexOf(' ')) : teacher,
                            bet,
                        null,
                            teacherTableCollection.TotalHours,
                           teacherTableCollection.AutumnHours,
                           teacherTableCollection.SpringHours,
                            null)
                            );

                    }
                }
                string tabname = prefix + "Итого";
                foreach(ExcelTotal list in totallist)
                {
                    list.DifferenceCalc();
                }
                TablesCollections.Add(new TableCollection(tabname, totallist));
                TablesCollections.SortTablesCollection();
                UpdateListBoxItemsSource();
            }
        }

        #endregion



        /// <summary>
        /// Вкладка Преподаватели
        /// </summary>

        //Открытие окна со списком преподавателей
        #region Show Teachers Window
        private RelayCommand _showTeachersWindowCommand;

        public ICommand ShowTeachersWindowCommand
        {
            get { return _showTeachersWindowCommand ?? (_showTeachersWindowCommand = new RelayCommand(ShowTeachersWindow)); }
        }

        private void ShowTeachersWindow(object obj)
        {
            var techerswindow = obj as Window;

            TeachersWindow teacherlist = new TeachersWindow();
            teacherlist.Owner = techerswindow;
            teacherlist.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            teacherlist.ShowDialog();
        }

        #endregion

        //Бланк нагрузки
        #region CalcReport
        private RelayCommand _loadCalcReport;
        private ReportViewModel loadCalcVM = new ReportViewModel();
        public ICommand LoadCalcReport
        {
            get { return _loadCalcReport ?? (_loadCalcReport = new RelayCommand(CreateLoadCalcReport)); }
        }

        private void CreateLoadCalcReport(object obj)
        {
            _=loadCalcVM.CreateLoadCalcAsync(SelectedComboBoxIndex);
        }
        #endregion

        //ИП преподавателей
        #region Individual Plan report
        private RelayCommand _individualPlanReport;

        public ICommand LoadIndividualPlanReport
        {
            get { return _individualPlanReport ?? (_individualPlanReport = new RelayCommand(CreateIndividualPlanReport)); }
        }

        private void CreateIndividualPlanReport(object obj)
        {
            _=loadCalcVM.CreateIndividualPlan(SelectedComboBoxIndex);
        }
        #endregion
    }
}
