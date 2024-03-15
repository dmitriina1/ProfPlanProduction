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

namespace ProfPlan.ViewModels
{
    internal class MainViewModel : ViewModel
    {
        public MainViewModel()
        {
            ExcelModel.UpdateSharedTeachers();
        }
        private string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Расчет нагрузки {DateTime.Today:dd-MM-yyyy}");
        private string filePath = "";
        private int Number = 1;
        private RelayCommand _loadDataCommand;
        private DataTableCollection tableCollection;
        public ICommand LoadDataCommand
        {
            get { return _loadDataCommand ?? (_loadDataCommand = new RelayCommand(LoadData)); }
        }
        private string GetExcelFilePath()
        {
            var openFileDialog = new OpenFileDialog() { Filter = "Excel Files|*.xls;*.xlsx" };

            return openFileDialog.ShowDialog() == true ? openFileDialog.FileName : null;
        }
        private DataSet ReadExcelData(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return null;
            directoryPath = filePath;
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
            string tabname = "";
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
                        MessageBox.Show($"Error adding data: {ex.Message}");
                    }
                }
            }
            else if (table.TableName.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
            {
                ProcessTotalTable(table, list);
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

        //Сохранение 
        private RelayCommand _saveDataToExcel;
        public ICommand SaveDataCommand
        {
            get { return _saveDataToExcel ?? (_saveDataToExcel = new RelayCommand(SaveToExcel)); }
        }
        private RelayCommand _saveDataToExcelAs;
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
        private void SaveToExcel(object parameter)
        {
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
            if (directoryPath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            {
                SaveToExcelAs();
            }
            else if (filePath != "")
            {
                workbook.SaveAs(filePath);
            }
            else
            {
                SaveToExcelAs();
            }
        }

        //Выход
        private RelayCommand _exitCommand;

        public ICommand ExitCommand
        {
            get { return _exitCommand ?? (_exitCommand = new RelayCommand(ExitFromApp)); }
        }
        private void ExitFromApp(object parameter)
        {
            Application.Current.Dispatcher.Invoke(() => Application.Current.Shutdown());
        }

        //Создать
        private RelayCommand _createCommand;

        public ICommand CreateCommand
        {
            get { return _createCommand ?? (_createCommand = new RelayCommand(CreateBaseTableCollection)); }
        }
        private void CreateBaseTableCollection(object parameter)
        {
            if(SelectedComboBoxIndex == 0 && TablesCollections.GetTableByName("ПИиИС",SelectedComboBoxIndex) == false)
            {
                TablesCollections.Add(new TableCollection() { Tablename = "П_ПИиИС" });
            }
            else if (SelectedComboBoxIndex == 1 && TablesCollections.GetTableByName("ПИиИС", SelectedComboBoxIndex) == false)
            {
                TablesCollections.Add(new TableCollection() { Tablename = "Ф_ПИиИС" });
            }
            UpdateListBoxItemsSource();
        }

        //Таблица
        //Очистить таблицу
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

        //Перенести преподавателей
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
            SelectedTable = null;
            UpdateListBoxItemsSource();
        }

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

        //Список преподавателей
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

        //Отчеты
        private RelayCommand _loadCalcReport;
        public ICommand LoadCalcReport
        {
            get { return _loadCalcReport ?? (_loadCalcReport = new RelayCommand(CreateLoadCalcReport)); }
        }

        private void CreateLoadCalcReport(object obj)
        {
            ReportViewModel loadCalcVM = new ReportViewModel();
            loadCalcVM.SumAllTeachersTables();
        }
    }
}
