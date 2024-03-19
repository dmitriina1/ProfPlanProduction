using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ProfPlan.Models;
using ProfPlan.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProfPlan.ViewModels
{
    internal class ReportViewModel : ViewModel
    {
        public ObservableCollection<TableCollection> TablesCollectionTeacherSum { get; set; }
        private ObservableCollection<TableCollection> TablesCollectionTeacherSumList { get; set; }
        private bool wasCalc = false;

        public void CreateLoadCalc(int index)
        {
            SumAllTeachersTables(index);
            wasCalc = true;
            SaveToExcel(TablesCollectionTeacherSum);
        }
        private void SumAllTeachersTables(int index)
        {
            TablesCollectionTeacherSum = new ObservableCollection<TableCollection>();

            TablesCollectionTeacherSumList = new ObservableCollection<TableCollection>();

            foreach (var tableCollection in TablesCollections.GetTablesCollection())
            {
                if (tableCollection.Tablename.IndexOf("ПИиИС", StringComparison.OrdinalIgnoreCase) == -1 && tableCollection.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) == -1 && tableCollection.Tablename.IndexOf("Незаполненные", StringComparison.OrdinalIgnoreCase) == -1 && tableCollection.Tablename.IndexOf("Доп", StringComparison.OrdinalIgnoreCase) == -1)
                {
                    double? bet = null;
                    double? betPercent = null;
                    double? totalHours = null;
                    double? autumnHours = null;
                    double? springHours = null;
                    string prefix;
                        if (index == 0)
                        {
                            prefix = "П_";
                        }
                        else
                        {
                            prefix = "Ф_";
                        }

                    var tablesCollection = index == 0 ? TablesCollections.GetTablesCollectionWithP() : TablesCollections.GetTablesCollectionWithF();

                    foreach (var tableCol in tablesCollection)
                    {
                        if (tableCol.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            foreach (ExcelTotal exRow in tableCol)
                            {
                                if (prefix + exRow.Teacher == tableCollection.Tablename)
                                {
                                    bet = exRow.Bet;
                                    betPercent = exRow.BetPercent;
                                    totalHours = exRow.TotalHours;
                                    autumnHours = exRow.AutumnHours;
                                    springHours = exRow.SpringHours;
                                    break;
                                }
                            }
                            break;
                        }
                    }
                    TableCollection sumTableCollection;
                    TableCollection sumTableCollectionTwo;
                    //Сумма колонок для Итого
                    ExcelModel sumOdd = CalculateSum(tableCollection, "нечет");
                    ObservableCollection<ExcelModel> sumOddList = TotalSemesterCalculation(tableCollection, "нечет");

                    ExcelModel sumEven = CalculateSum(tableCollection, "чет");
                    ObservableCollection<ExcelModel> sumEvenList = TotalSemesterCalculation(tableCollection, "чет");

                    double? autumnIndex = autumnHours/totalHours;
                    double? springIndex = springHours/totalHours;

                    if (betPercent == 1 || betPercent == null)
                    {
                        sumTableCollection = new TableCollection($"{tableCollection.Tablename}");
                        
                        if (bet!=null)
                        {
                            ProcessBet(sumTableCollection, sumOddList, sumEvenList, bet, betPercent, autumnIndex, springIndex);
                            TablesCollectionTeacherSum.Add(sumTableCollection);
                        }
                       
                    }
                    else
                    {
                        if (bet!=null)
                        {
                            if (betPercent > 1)
                            {
                                sumTableCollection = new TableCollection($"{tableCollection.Tablename}");
                                sumTableCollectionTwo = new TableCollection($"{tableCollection.Tablename} {betPercent - 1}");

                                ProcessBet(sumTableCollection, sumOddList, sumEvenList, bet, betPercent, autumnIndex, springIndex);
                                TablesCollectionTeacherSum.Add(sumTableCollection);

                                ProcessBet(sumTableCollectionTwo, sumOddList, sumEvenList, bet*(betPercent - 1), betPercent, autumnIndex, springIndex);
                                TablesCollectionTeacherSum.Add(sumTableCollectionTwo);
                            }
                            else
                            {
                                    sumTableCollection = new TableCollection($"{tableCollection.Tablename} {betPercent}");

                                    ProcessBet(sumTableCollection, sumOddList, sumEvenList, bet, betPercent, autumnIndex, springIndex);
                                    TablesCollectionTeacherSum.Add(sumTableCollection);
                                
                            }
                        }
                        
                    }                    


                }

            }
            
        }


        private TableCollection ProcessBet(TableCollection sumTableCollection, ObservableCollection<ExcelModel> sumOddList, ObservableCollection<ExcelModel> sumEvenList, double? bet, double? betPercent, double? autumnIndex = null, double? springIndex = null)
        {
            TableCollection sumOddListOneBet = new TableCollection();
            TableCollection sumEvenListOneBet = new TableCollection();

            TableCollection ListForIPPlan = new TableCollection($"{sumTableCollection.Tablename}");

            double? sum = 0, betValue, min = null, dif;
            
            betValue = bet * autumnIndex;
            
            if (betPercent<1)
            {
                betValue = bet * autumnIndex * betPercent;
            }
            foreach (ExcelModel excelModel in sumOddList)
            {
                if (betValue>sum)
                {
                    dif=betValue - sum;
                    if (dif > excelModel.SumProperties())
                    {
                        sum+=excelModel.SumProperties();
                        sumOddListOneBet.ExcelData.Add(excelModel);
                        ListForIPPlan.ExcelData.Add(excelModel);
                    }
                    else if(min>excelModel.SumProperties() && sumOddList.IndexOf(excelModel) != sumOddList.Count - 1)
                    {
                        min = excelModel.SumProperties();
                    }
                    else if(sumOddList.IndexOf(excelModel) == sumOddList.Count - 1)
                    {
                        if (min == null)
                            min = 0;
                        sum+=min;
                        sumOddListOneBet.ExcelData.Add(excelModel);
                        ListForIPPlan.ExcelData.Add(excelModel);
                    }
                }
                else
                {
                    break;
                }
            }
            DeleteItemsFromObsCol(sumOddList, sumOddListOneBet.ExcelData.Count-1);
            
            sum = 0;
            min = null;
            betValue = bet * springIndex;
            if (betPercent<1)
            {
                betValue = bet * springIndex * betPercent;
            }
            foreach (ExcelModel excelModel in sumEvenList)
            {
                if (betValue>sum)
                {
                    dif=betValue - sum;
                    if (dif > excelModel.SumProperties())
                    {
                        sum += excelModel.SumProperties();
                        sumEvenListOneBet.ExcelData.Add(excelModel);
                        ListForIPPlan.ExcelData.Add(excelModel);
                    }
                    else if (min>excelModel.SumProperties() && sumOddList.IndexOf(excelModel) != sumOddList.Count - 1)
                    {
                        min = excelModel.SumProperties();
                    }
                    else if (sumEvenList.IndexOf(excelModel) == sumEvenList.Count - 1)
                    {
                        if (min == null)
                            min = 0;
                        sum+=min;
                        sumEvenListOneBet.ExcelData.Add(excelModel);
                        ListForIPPlan.ExcelData.Add(excelModel);
                    }
                }
                else
                {
                    break;
                }
            }

            DeleteItemsFromObsCol(sumEvenList, sumEvenListOneBet.ExcelData.Count-1);
            ExcelModel sumOddOneBet = CalculateSum(sumOddListOneBet, "нечет");
            ExcelModel sumEvenOneBet = CalculateSum(sumEvenListOneBet, "чет");
            sumTableCollection.ExcelData.Add(sumOddOneBet);
            sumTableCollection.ExcelData.Add(sumEvenOneBet);
            TablesCollectionTeacherSumList.Add(ListForIPPlan);
            return sumTableCollection;
        }


        private ObservableCollection<ExcelModel> TotalSemesterCalculation(TableCollection tableCollection, string term)
        {
            ObservableCollection<ExcelModel> ex = new ObservableCollection<ExcelModel>();
            foreach (var excelModel in tableCollection.ExcelData.OfType<ExcelModel>().Where(x => x.Term != null && x.Term.Equals(term, StringComparison.OrdinalIgnoreCase)))
            {
                ex.Add(excelModel);
            }
            return ex;
        }

        private ExcelModel CalculateSum(TableCollection tableCollection, string term)
        {
            var sumModel = new ExcelModel( 0, "", "", term, "", "", null, "", "", null, null, null,
                null, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

            foreach (var excelModel in tableCollection.ExcelData.OfType<ExcelModel>().Where(x => x.Term != null && x.Term.Equals(term, StringComparison.OrdinalIgnoreCase)))
            {
                if (excelModel.Lectures!=null)
                    sumModel.Lectures += excelModel.Lectures;
                if (excelModel.Consultations != null)
                    sumModel.Consultations += excelModel.Consultations;
                if (excelModel.Laboratory != null)
                    sumModel.Laboratory += excelModel.Laboratory;
                if (excelModel.Practices != null)
                    sumModel.Practices += excelModel.Practices;
                if (excelModel.Tests != null)
                    sumModel.Tests += excelModel.Tests;
                if (excelModel.Exams != null)
                    sumModel.Exams += excelModel.Exams;
                if (excelModel.CourseProjects != null)
                    sumModel.CourseProjects += excelModel.CourseProjects;
                if (excelModel.CourseWorks != null)
                    sumModel.CourseWorks += excelModel.CourseWorks;
                if (excelModel.Diploma != null)
                    sumModel.Diploma += excelModel.Diploma;
                if (excelModel.RGZ != null)
                    sumModel.RGZ += excelModel.RGZ;
                if (excelModel.GEKAndGAK != null)
                    sumModel.GEKAndGAK += excelModel.GEKAndGAK;
                if (excelModel.ReviewDiploma != null)
                    sumModel.ReviewDiploma += excelModel.ReviewDiploma;
                if (excelModel.Other != null)
                    sumModel.Other += excelModel.Other;
            }
            if (sumModel.Lectures == 0)
                sumModel.Lectures = null;
            if (sumModel.Consultations == 0)
                sumModel.Consultations = null;
            if (sumModel.Laboratory == 0)
                sumModel.Laboratory = null;
            if (sumModel.Practices == 0)
                sumModel.Practices = null;
            if (sumModel.Tests == 0)
                sumModel.Tests = null;
            if (sumModel.Exams == 0)
                sumModel.Exams = null;
            if (sumModel.CourseProjects == 0)
                sumModel.CourseProjects = null;
            if (sumModel.CourseWorks == 0)
                sumModel.CourseWorks = null;
            if (sumModel.Diploma == 0)
                sumModel.Diploma = null;
            if (sumModel.RGZ == 0)
                sumModel.RGZ = null;
            if (sumModel.GEKAndGAK == 0)
                sumModel.GEKAndGAK = null;
            if (sumModel.ReviewDiploma == 0)
                sumModel.ReviewDiploma = null;
            if (sumModel.Other == 0)
                sumModel.Other = null;

            return sumModel;
        }


        private void DeleteItemsFromObsCol(ObservableCollection<ExcelModel> collection, int indexToRemoveUpTo)
        {
            for (int i = 0; i <= indexToRemoveUpTo && collection.Count > 0; i++)
            {
                collection.RemoveAt(0);
            }
        }


        private string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Бланк нагрузки {DateTime.Today:dd-MM-yyyy}");

        private string GetSaveFilePath()
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.FileName = $"Бланк_Нагрузки {DateTime.Today:dd-MM-yyyy}.xlsx";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                return saveFileDialog.FileName;

            return null;
        }

        private string CreateTeachersNameForForm(string teacherName)
        {
            if (teacherName.StartsWith("П_") || teacherName.StartsWith("Ф_"))
            {
                teacherName = teacherName.Substring(2);
            }
            foreach (var teach in TeachersManager.GetTeachers())
            {
                if (Regex.Replace(teacherName.Trim(), @"\s+", " ").Split(' ')[0] == Regex.Replace(teach.LastName.Trim(), @"\s+", " "))
                {
                    if (Regex.Replace(teacherName.Trim(), @"\s+", " ").Split(' ').Length > 1)
                    {
                        teacherName = teach.LastName + "\n" + teach.FirstName + "\n" + teach.MiddleName + "\n" + teach.AcademicDegree + " " +
                        teach.Position +"\n" + (Regex.Replace(teacherName.Trim(), @"\s+", " ").Split(' ').Length > 1 ? Regex.Replace(teacherName.Trim(), @"\s+", " ").Split(' ')[1] : "") + " ставки";
                    }
                    else if (teach.AcademicDegree!="" && teach.Position!="")
                    {
                        teacherName = teach.LastName + "\n" + teach.FirstName + "\n" + teach.MiddleName + "\n" + teach.AcademicDegree + " " + teach.Position;

                    }
                    else
                    {
                        teacherName = teach.LastName + "\n" + teach.FirstName + "\n" + teach.MiddleName;

                    }

                }
            }
            return teacherName;
        }

        public void SaveToExcel(ObservableCollection<TableCollection> tablesCollection)
        {
            directoryPath = GetSaveFilePath();
            if (string.IsNullOrEmpty(directoryPath))
                return;

            using (var workbook = new XLWorkbook())
            {
                var fworksheet = workbook.Worksheets.Add("Бланк нагрузки");
                int frow = 3;
                // Добавление заголовков
                int columnNumber = 1;

                fworksheet.Cell(frow, columnNumber++).Value = "Teacher";

                List<string> propertyNames = new List<string>();

                foreach (var propertyInfo in typeof(ExcelModel).GetProperties())
                {
                    if (propertyInfo.Name == "Lectures" || propertyInfo.Name == "Consultations" || propertyInfo.Name == "Laboratory" || propertyInfo.Name == "Practices" || propertyInfo.Name == "Tests" || propertyInfo.Name == "Exams" || propertyInfo.Name == "CourseProjects" || propertyInfo.Name == "CourseWorks" || propertyInfo.Name == "Diploma" || propertyInfo.Name == "RGZ" || propertyInfo.Name == "GEKAndGAK" || propertyInfo.Name == "ReviewDiploma" || propertyInfo.Name == "Other")
                    {
                        fworksheet.Cell(frow, columnNumber).Value = propertyInfo.Name;
                        propertyNames.Add(propertyInfo.Name);
                        columnNumber++;
                    }
                }

                fworksheet.Cell(3, columnNumber).Value = "TotalSemester";

                // Заполнение данных - первые элементы
                int rowNumber = 4;
                foreach (var tableCollection in tablesCollection)
                {
                    string teacherName = CreateTeachersNameForForm(tableCollection.Tablename);
                    
                    if (tableCollection.ExcelData.Count >= 1)
                    {
                        var excelModel = tableCollection.ExcelData[0];
                        columnNumber = 1;
                        fworksheet.Cell(rowNumber, columnNumber++).Value = teacherName;

                        // Сумма колонок
                        double totalSemester = propertyNames.Sum(propertyName => Convert.ToDouble(typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null) ?? 0));
                        foreach (var propertyName in propertyNames)
                        {
                            var value = typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null);
                            fworksheet.Cell(rowNumber, columnNumber++).Value = value != null ? value.ToString() : "";
                        }

                        fworksheet.Cell(rowNumber, columnNumber).Value = totalSemester.ToString().Replace(",",".");

                        rowNumber++;
                    }
                }

                //// Дублирую колонки
                int headerRow = 3;
               
                rowNumber=4;

                //columnNumber = 1;
                int cnum = 16;
                foreach (var propertyName in propertyNames)
                {
                    fworksheet.Cell(headerRow, cnum++).Value = propertyName;
                }
                fworksheet.Cell(headerRow, cnum).Value = "TotalSemester";

                // Заполнение данных - вторые элементы
                foreach (var tableCollection in tablesCollection)
                {

                    if (tableCollection.ExcelData.Count == 2)
                    {
                        var excelModel = tableCollection.ExcelData[1];

                        cnum = 16;

                        // Сумма
                        double totalSemester = propertyNames.Sum(propertyName => Convert.ToDouble(typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null) ?? 0));
                        foreach (var propertyName in propertyNames)
                        {
                            var value = typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null);
                            fworksheet.Cell(rowNumber, cnum++).Value = value != null ? value.ToString() : "";
                        }


                        fworksheet.Cell(rowNumber, cnum).Value = totalSemester.ToString().Replace(",", ".");

                        rowNumber++;
                    }
                }
                AdjustWorksheetLayout(fworksheet, frow, workbook);
                // Сохранение в файл
                workbook.SaveAs(directoryPath);
            }
        }
        private void AdjustWorksheetLayout(IXLWorksheet fworksheet, int frow, XLWorkbook workbook)
        {
            SwapAndInsertColumns(fworksheet, frow);

            // Задаем заголовок для нового столбца
            fworksheet.Cell(frow, 42).Value = "ИТОГО ЗА ГОД";
            // Заполняем значениями новый столбец на основе данных из других столбцов
            for (int row = 4; row <= fworksheet.RowsUsed().Count()+2; row++)
            {
                var value21 = fworksheet.Cell(row, 21).Value.ToString().ToNullable<double>();
                var value41 = fworksheet.Cell(row, 41).Value.ToString().ToNullable<double>();
                fworksheet.Cell(row, 42).Value = (value21 + value41).ToString().Replace(",", ".");
            }
            frow = 1;
            //sworksheet.Range(frow, 1, frow, 3).Merge();
            fworksheet.Range(frow, 1, frow, 3).Merge();
            fworksheet.Cell(frow, 1).Value = "Первое полугодие";
            fworksheet.Range(frow, 22, frow, 27).Merge();
            fworksheet.Cell(frow, 22).Value = "Второе полугодие";

            fworksheet.Range(fworksheet.Cell(2, 8), fworksheet.Cell(2, 14)).Merge();
            fworksheet.Cell(2, 8).Value = "Руководство";
            fworksheet.Range(fworksheet.Cell(2, 28), fworksheet.Cell(2, 34)).Merge();
            fworksheet.Cell(2, 28).Value = "Руководство";

            List<string> newPropertyNames = GetPropertyNamesForColumns();

            for (int col = 1; col <= 43; col++)
            {
                if ((col < 8 || col > 14) && (col < 28 || col > 34)) // Проверяем, что колонка не входит в диапазоны 8-14 и 28-34
                {
                    fworksheet.Range(fworksheet.Cell(2, col), fworksheet.Cell(3, col)).Merge();
                }
            }
            for (int i = 1; i < newPropertyNames.Count; i++)
            {
                if (i<7 ||i>13)
                {
                    fworksheet.Cell(2, i + 1 + 20).Value = newPropertyNames[i];
                    fworksheet.Cell(2, i + 1).Value = newPropertyNames[i];
                }

            }
            SetStyleForWorksheet(fworksheet, workbook);
        }
        private void SetStyleForWorksheet(IXLWorksheet fworksheet, XLWorkbook workbook)
        {
            fworksheet.Cell(2, 42).Value = "ИТОГО ЗА ГОД";

            var styleArial6 = workbook.Style;
            styleArial6.Alignment.TextRotation = 90;
            styleArial6.Alignment.WrapText = true;
            for (int row = 2; row <= fworksheet.RowsUsed().Count()+1; row++)
            {
                for (int col = 2; col <= fworksheet.ColumnsUsed().Count(); col++)
                {
                    fworksheet.Cell(row, col).Style = styleArial6;
                    if (double.TryParse(fworksheet.Cell(row, col).Value.ToString(), out double number))
                    {
                        // Проверяем, имеет ли число дробную часть
                        if (number != Math.Floor(number))
                        {
                            // Если число не является целым, устанавливаем явный формат
                            fworksheet.Cell(row, col).Style.NumberFormat.Format = "0.##";
                        }
                    }
                }
            }
            fworksheet.Column(1).Style.Alignment.WrapText = true;
            fworksheet.Cell(2, 8).Style = workbook.Style.Alignment.SetTextRotation(0);
            fworksheet.Cell(2, 28).Style = workbook.Style.Alignment.SetTextRotation(0);
            fworksheet.Columns().AdjustToContents();
            fworksheet.Rows(2, 3).AdjustToContents();
        }
        private List<string> GetPropertyNamesForColumns()
        {
           return new List<string>
                {
                    "Преподаватель", "Чтение лекций", "Консультации", "Лабораторные работы",
                    "Практические занятия", "Зачеты", "Экзамены", "Курсовыми проектами",
                    "Курсовыми работами", "Дипломными работами", "Учебной практикой", "Произв. практикой",
                    "УИРС", "Аспирантами и соискат.", "РГР", "Консультации для заочников",
                    "Рецензирование контр. Работ заочников", "ГЭК",
                    "Проверка контрольных работ", "Другие виды работ", "ИТОГО ЗА СЕМЕСТР"
                };
        }

        private void SwapAndInsertColumns(IXLWorksheet fworksheet, int frow)
        {
            SwapColumns(fworksheet, 3, 5);
            SwapColumns(fworksheet, 8, 9);
            SwapColumns(fworksheet, 10, 11);
            SwapColumns(fworksheet, 11, 12);
            SwapColumns(fworksheet, 17, 19);
            SwapColumns(fworksheet, 22, 23);
            SwapColumns(fworksheet, 24, 25);
            SwapColumns(fworksheet, 25, 26);
           
            fworksheet.Column(11).InsertColumnsBefore(4);
            fworksheet.Column(16).InsertColumnsBefore(2);
            fworksheet.Column(31).InsertColumnsBefore(4);
            fworksheet.Column(36).InsertColumnsBefore(2);
            List<string> newPropertyNames = GetPropertyNamesForColumns();


            fworksheet.Cell(frow, 1).Value="";
            for (int i = 1; i < newPropertyNames.Count; i++)
            {
                fworksheet.Cell(frow, i + 1 + 20).Value = newPropertyNames[i];
                fworksheet.Cell(frow, i + 1).Value = newPropertyNames[i];
            }
            if (fworksheet.Column(42) == null)
            {
                fworksheet.Column(41).InsertColumnsAfter(1);
            }
        }
        public static void SwapColumns(IXLWorksheet worksheet, int column1Index, int column2Index)
        {
            int startRow = worksheet.FirstRowUsed().RowNumber();
            int endRow = worksheet.LastRowUsed().RowNumber();

            for (int row = startRow; row <= endRow; row++)
            {
                var tempValue = worksheet.Cell(row, column1Index).Value;
                worksheet.Cell(row, column1Index).Value = worksheet.Cell(row, column2Index).Value;
                worksheet.Cell(row, column2Index).Value = tempValue;
            }
        }

        //ИП Преподователей

        public void CreateIndividualPlan(int index)
        {
            if(wasCalc == false)
            {
                SumAllTeachersTables(index);
            }
            var workbook = new XLWorkbook();
            foreach(TableCollection tab in TablesCollectionTeacherSumList)
            {
                var worksheet = workbook.Worksheets.Add(tab.Tablename);
                worksheet.Cell(1, 1).Value = "Дисциплина";
                worksheet.Cell(1, 2).Value = "Группа";
                worksheet.Cell(1, 3).Value = "Институт";
                worksheet.Cell(1, 4).Value = "Часы";
                int row = 2;
                foreach (ExcelModel excel in tab.ExcelData)
                {
                    
                        worksheet.Cell(row, 1).Value = excel.Discipline;
                        worksheet.Cell(row, 2).Value = excel.Group;
                        worksheet.Cell(row, 3).Value = excel.SubGroup;
                        worksheet.Cell(row, 4).Value = excel.Institute;
                        worksheet.Cell(row, 5).Value = excel.Total;
                        row++;
                    
                }
            }
            directoryPath = GetSaveFilePath();
            if (string.IsNullOrEmpty(directoryPath))
                return;
            workbook.SaveAs(directoryPath);
        }

        private static bool IsValidGroup(string group)
        {
            // Проверяем, что строка не пустая и содержит хотя бы одну букву
            return !string.IsNullOrWhiteSpace(group) && group.Any(c => char.IsLetter(c));
        }

    }
}
