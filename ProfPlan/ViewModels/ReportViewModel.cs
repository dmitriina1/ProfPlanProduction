using ClosedXML.Excel;
using ProfPlan.Models;
using ProfPlan.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProfPlan.ViewModels
{
    internal class ReportViewModel : ViewModel
    {
        public ObservableCollection<TableCollection> TablesCollectionTeacherSum { get; set; }

        //private int? _orientation = 0;
        //public int? Orientation
        //{
        //    get { return _orientation; }
        //    set
        //    {
        //        if (_orientation != value)
        //        {
        //            _orientation = value;
        //            OnPropertyChanged(nameof(Orientation));
        //        }
        //    }
        //}
        public void SumAllTeachersTables(int index)
        {
            TablesCollectionTeacherSum = new ObservableCollection<TableCollection>();
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
                    foreach (var tableCol in TablesCollections.GetTablesCollection())
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
                                }
                            }
                        }
                    }
                    TableCollection sumTableCollection;
                    TableCollection sumTableCollectionTwo;
                    //Сумма колонок для Итого
                    ExcelModel sumOdd = CalculateSum(tableCollection, "нечет");
                    //Коллекция записей с Term = нечет
                    ObservableCollection<ExcelModel> sumOddList = TotalSemesterCalculation(tableCollection, "нечет");
                    //Total сумма
                    //double sumod = SumObsColExModel(sumOdd);

                    ExcelModel sumEven = CalculateSum(tableCollection, "чет");
                    ObservableCollection<ExcelModel> sumEvenList = TotalSemesterCalculation(tableCollection, "чет");
                    //double sumev = SumObsColExModel(sumEven);
                    double sum;
                    double? autumnIndex = autumnHours/totalHours;
                    double? springIndex = springHours/totalHours;
                    if (betPercent == 1 || betPercent == null)
                    {
                        sumTableCollection = new TableCollection($"{tableCollection.Tablename}");
                        
                        if (bet!=null)
                        {
                            TableCollection sumOddListOneBet = new TableCollection();
                            TableCollection sumEvenListOneBet = new TableCollection();
                            sum = 0;
                            foreach(ExcelModel excelModel in sumOddList)
                            {
                                if (bet>sum)
                                {
                                    sum+=excelModel.SumProperties();
                                    sumOddListOneBet.ExcelData.Add(excelModel);

                                }
                                else
                                {
                                    DeleteItemsFromObsCol(sumOddList, sumOddListOneBet.ExcelData.Count-1);
                                    break;
                                }
                            }
                            sum = 0;
                            foreach (ExcelModel excelModel in sumEvenList)
                            {
                                if (bet > sum)
                                {
                                    sum += excelModel.SumProperties();
                                    sumEvenListOneBet.ExcelData.Add(excelModel);
                                }
                                else
                                {
                                    DeleteItemsFromObsCol(sumEvenList, sumEvenListOneBet.ExcelData.Count-1);
                                    break;
                                }
                            }
                            ExcelModel sumOddOneBet = CalculateSum(sumOddListOneBet, "нечет");
                            ExcelModel sumEvenOneBet = CalculateSum(sumEvenListOneBet, "чет");
                            sumTableCollection.ExcelData.Add(sumOddOneBet);
                            sumTableCollection.ExcelData.Add(sumEvenOneBet);
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
                                TableCollection sumOddListOneBet = new TableCollection();
                                TableCollection sumOddListTwoBet = new TableCollection();
                                TableCollection sumEvenListOneBet = new TableCollection();
                                TableCollection sumEvenListTwoBet = new TableCollection();

                                sum = 0;
                                foreach (ExcelModel excelModel in sumOddList)
                                {
                                    if (bet*autumnIndex > sum)
                                    {
                                        sum += excelModel.SumProperties();
                                        sumOddListOneBet.ExcelData.Add(excelModel);
                                    }
                                    else
                                    {
                                        DeleteItemsFromObsCol(sumOddList, sumOddListOneBet.ExcelData.Count-1);
                                        break;
                                    }
                                }
                                sum = 0;
                                foreach (ExcelModel excelModel in sumOddList)
                                {
                                    if ((bet*autumnIndex * (betPercent - 1)) > sum)
                                    {
                                        sum += excelModel.SumProperties();
                                        sumOddListTwoBet.ExcelData.Add(excelModel);
                                    }
                                    else
                                    {
                                        DeleteItemsFromObsCol(sumOddList, sumOddListTwoBet.ExcelData.Count-1);
                                        break;
                                    }
                                }

                                sum = 0;
                                foreach (ExcelModel excelModel in sumEvenList)
                                {
                                    if (bet*springIndex > sum)
                                    {
                                        sum += excelModel.SumProperties();
                                        sumEvenListOneBet.ExcelData.Add(excelModel);
                                    }
                                    else
                                    {
                                        DeleteItemsFromObsCol(sumEvenList, sumEvenListOneBet.ExcelData.Count-1);
                                        break;
                                    }
                                }
                                sum = 0;
                                foreach (ExcelModel excelModel in sumEvenList)
                                {
                                    if ((bet*springIndex * (betPercent - 1)) > sum)
                                    {
                                        sum += excelModel.SumProperties();
                                        sumEvenListTwoBet.ExcelData.Add(excelModel);
                                    }
                                    else
                                    {
                                        DeleteItemsFromObsCol(sumEvenList, sumEvenListTwoBet.ExcelData.Count-1);
                                        break;
                                    }
                                }
                                ExcelModel sumOddOneBet = CalculateSum(sumOddListOneBet, "нечет");
                                ExcelModel sumOddTwoBet = CalculateSum(sumOddListTwoBet, "нечет");
                                ExcelModel sumEvenOneBet = CalculateSum(sumEvenListOneBet, "чет");
                                ExcelModel sumEvenTwoBet = CalculateSum(sumEvenListTwoBet, "чет");
                                sumTableCollection.ExcelData.Add(sumOddOneBet);
                                sumTableCollection.ExcelData.Add(sumEvenOneBet);
                                TablesCollectionTeacherSum.Add(sumTableCollection);

                                sumTableCollectionTwo.ExcelData.Add(sumOddTwoBet);
                                sumTableCollectionTwo.ExcelData.Add(sumEvenTwoBet);
                                TablesCollectionTeacherSum.Add(sumTableCollectionTwo);
                            }
                        }
                        else
                        {
                            if (bet!=null)
                            {
                                sumTableCollection = new TableCollection($"{tableCollection.Tablename}");
                                TableCollection sumListOneBet = new TableCollection();
                                sum = 0;
                                foreach (ExcelModel excelModel in sumOddList)
                                {
                                    if ((bet*autumnIndex * betPercent) > sum)
                                    {
                                        sum += excelModel.SumProperties();
                                        sumListOneBet.ExcelData.Add(excelModel);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                sum = 0;
                                foreach (ExcelModel excelModel in sumEvenList)
                                {
                                    if ((bet*springIndex * betPercent) > sum)
                                    {
                                        sum += excelModel.SumProperties();
                                        sumListOneBet.ExcelData.Add(excelModel);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                ExcelModel sumOddBet = CalculateSum(sumListOneBet, "нечет");
                                ExcelModel sumEvenBet = CalculateSum(sumListOneBet, "чет");
                                sumTableCollection.ExcelData.Add(sumOddBet);
                                sumTableCollection.ExcelData.Add(sumEvenBet);
                                TablesCollectionTeacherSum.Add(sumTableCollection);
                            }
                        }
                    }



                }
            }
            SaveToExcel(TablesCollectionTeacherSum);
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

        public void SaveToExcel(ObservableCollection<TableCollection> tablesCollection)
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.FileName = $"Бланк_Нагрузки {DateTime.Today:dd-MM-yyyy}.xlsx";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                directoryPath = saveFileDialog.FileName;
            }
            else
            {
                return;
            }
            using (var workbook = new XLWorkbook())
            {
                var fworksheet = workbook.Worksheets.Add("Первое полугодие");
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
                    string teacherName = tableCollection.Tablename;

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

                //rowNumber += 6;
                //var sworksheet = workbook.Worksheets.Add("Второе полугодие");
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
                    string teacherName = tableCollection.Tablename;

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
                SwapColumns(fworksheet, 3, 5);
                SwapColumns(fworksheet, 8, 9);
                SwapColumns(fworksheet, 10, 11);
                SwapColumns(fworksheet, 11, 12);
                SwapColumns(fworksheet, 17, 19);
                SwapColumns(fworksheet, 22, 23);
                SwapColumns(fworksheet, 24, 25);
                SwapColumns(fworksheet, 25, 26);
                List<string> newPropertyNames = new List<string>
                {
                    "Преподаватель", "Чтение лекций", "Консультации", "Лабораторные работы",
                    "Практические занятия", "Зачеты", "Экзамены", "Курсовыми проектами",
                    "Курсовыми работами", "Дипломными работами", "Учебной практикой", "Произв. практикой",
                    "УИРС", "Аспирантами и соискат.", "РГР", "Консультации для заочников",
                    "Рецензирование контр. Работ заочников", "ГЭК",
                    "Проверка контрольных работ", "Другие виды работ", "ИТОГО ЗА СЕМЕСТР"
                };
                fworksheet.Column(11).InsertColumnsBefore(4);
                fworksheet.Column(16).InsertColumnsBefore(2);
                fworksheet.Column(31).InsertColumnsBefore(4);
                fworksheet.Column(36).InsertColumnsBefore(2);

                //for(int i = 1; i<fworksheet.ColumnsUsed().Count()+1; i++)
                //{
                //    fworksheet.Range(fworksheet.Cell(2, i), fworksheet.Cell(3, i)).Merge();
                //}
                //fworksheet.Range(fworksheet.Cell(2, 1), fworksheet.Cell(3, 1)).Merge();
                //fworksheet.Range(fworksheet.Cell(2, 2), fworksheet.Cell(3, 2)).Merge();
                //fworksheet.Range(fworksheet.Cell(2, 3), fworksheet.Cell(3, 3)).Merge();
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

                // Задаем заголовок для нового столбца
                fworksheet.Cell(frow, 42).Value = "ИТОГО ЗА ГОД";
                // Заполняем значениями новый столбец на основе данных из других столбцов
                for (int row = 4; row <= fworksheet.RowsUsed().Count()+2; row++)
                {
                    var value21 = fworksheet.Cell(row, 21).Value.ToString().ToNullable<double>();
                    var value41 = fworksheet.Cell(row, 41).Value.ToString().ToNullable<double>();
                    fworksheet.Cell(row, 42).Value = (value21 + value41).ToString().Replace(",",".");
                }
                frow = 1;
                //sworksheet.Range(frow, 1, frow, 3).Merge();
                fworksheet.Range(frow, 1, frow, 3).Merge();
                fworksheet.Cell(frow, 1).Value = "Первое полугодие";
                fworksheet.Range(frow, 22, frow, 27).Merge();
                fworksheet.Cell(frow, 22).Value = "Второе полугодие";


                var styleArial6 = workbook.Style;
                styleArial6.Alignment.TextRotation = 90;
                for (int row = 3; row <= fworksheet.RowsUsed().Count()+1; row++)
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
               
                fworksheet.Columns().AdjustToContents();
                fworksheet.Rows(2,3).AdjustToContents();

                // Сохранение в файл
                workbook.SaveAs(directoryPath);
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
    }
}
