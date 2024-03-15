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

namespace ProfPlan.ViewModels
{
    internal class ReportViewModel : ViewModel
    {
        public ObservableCollection<TableCollection> TablesCollectionTeacherSum { get; set; }

        public void SumAllTeachersTables()
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
                    foreach (var tableCol in TablesCollections.GetTablesCollection())
                    {
                        if (tableCol.Tablename.IndexOf("Итого", StringComparison.OrdinalIgnoreCase) != -1)
                        {
                            foreach (ExcelTotal exRow in tableCol)
                            {
                                if (exRow.Teacher.Split(' ').FirstOrDefault() == tableCollection.Tablename)
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

                    if (betPercent == 1 || betPercent == null)
                    {
                        sumTableCollection = new TableCollection($"{tableCollection.Tablename}");
                        sumTableCollection.ExcelData.Add(sumOdd);
                        sumTableCollection.ExcelData.Add(sumEven);
                        TablesCollectionTeacherSum.Add(sumTableCollection);
                    }
                    else
                    {
                        double sum;
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
                                if ((autumnHours/betPercent) > sum)
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
                                if ((autumnHours * (betPercent - 1)) > sum)
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
                                if ((springHours / betPercent) > sum)
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
                                if ((springHours * (betPercent - 1)) > sum)
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
                        else
                        {
                            sumTableCollection = new TableCollection($"{tableCollection.Tablename}");
                            TableCollection sumListOneBet = new TableCollection();
                            sum = 0;
                            foreach (ExcelModel excelModel in sumOddList)
                            {
                                if ((autumnHours * betPercent) > sum)
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
                                if ((springHours * betPercent) > sum)
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
                int frow = 1;
                // Добавление заголовков

                int columnNumber = 1;

                fworksheet.Range(frow, 1, frow, 3).Merge();
                fworksheet.Cell(frow, 1).Value = "Первое полугодие";
                frow++;

                fworksheet.Cell(2, columnNumber++).Value = "Teacher";

                List<string> propertyNames = new List<string>();

                foreach (var propertyInfo in typeof(ExcelModel).GetProperties())
                {
                    if (propertyInfo.Name == "Lectures" || propertyInfo.Name == "Consultations" || propertyInfo.Name == "Laboratory" || propertyInfo.Name == "Practices" || propertyInfo.Name == "Tests" || propertyInfo.Name == "Exams" || propertyInfo.Name == "CourseProjects" || propertyInfo.Name == "CourseWorks" || propertyInfo.Name == "Diploma" || propertyInfo.Name == "RGZ" || propertyInfo.Name == "GEKAndGAK" || propertyInfo.Name == "ReviewDiploma" || propertyInfo.Name == "Other")
                    {
                        fworksheet.Cell(2, columnNumber).Value = propertyInfo.Name;
                        propertyNames.Add(propertyInfo.Name);
                        columnNumber++;
                    }
                }


                fworksheet.Cell(2, columnNumber).Value = "TotalSemester";

                // Заполнение данных - первые элементы
                int rowNumber = 3;
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


                        fworksheet.Cell(rowNumber, columnNumber).Value = totalSemester;

                        rowNumber++;
                    }
                }


                rowNumber += 6;
                var sworksheet = workbook.Worksheets.Add("Второе полугодие");
                // Дублирую колонки
                int headerRow = 2;
                frow = 1;
                sworksheet.Range(frow, 1, frow, 3).Merge();
                sworksheet.Cell(frow, 1).Value = "Второе полугодие";
                frow++;
                rowNumber=3;

                columnNumber = 1;
                sworksheet.Cell(headerRow, columnNumber++).Value = "Teacher";
                foreach (var propertyName in propertyNames)
                {
                    sworksheet.Cell(headerRow, columnNumber++).Value = propertyName;
                }
                sworksheet.Cell(headerRow, columnNumber).Value = "TotalSemester";

                // Заполнение данных - вторые элементы
                foreach (var tableCollection in tablesCollection)
                {
                    string teacherName = tableCollection.Tablename;

                    if (tableCollection.ExcelData.Count == 2)
                    {
                        var excelModel = tableCollection.ExcelData[1];

                        columnNumber = 1;
                        sworksheet.Cell(rowNumber, columnNumber++).Value = teacherName;

                        // Сумма
                        double totalSemester = propertyNames.Sum(propertyName => Convert.ToDouble(typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null) ?? 0));
                        foreach (var propertyName in propertyNames)
                        {
                            var value = typeof(ExcelModel).GetProperty(propertyName)?.GetValue(excelModel, null);
                            sworksheet.Cell(rowNumber, columnNumber++).Value = value != null ? value.ToString() : "";
                        }


                        sworksheet.Cell(rowNumber, columnNumber).Value = totalSemester;

                        rowNumber++;
                    }
                }
                SwapColumns(fworksheet, 3, 5);
                SwapColumns(fworksheet, 8, 9);
                SwapColumns(fworksheet, 10, 11);
                SwapColumns(fworksheet, 11, 12);
                SwapColumns(sworksheet, 3, 5);
                SwapColumns(sworksheet, 8, 9);
                SwapColumns(sworksheet, 10, 11);
                SwapColumns(sworksheet, 11, 12);
                List<string> newPropertyNames = new List<string>
                {
                    "Преподаватель", "Чтение лекций", "Консультации", "Лабораторные работы",
                    "Практические занятия", "Зачеты", "Экзамены", "Курсовыми проектами",
                    "Курсовыми работами", "Дипломными работами", "РГР", "ГЭК",
                    "Проверка контрольных работ", "Другие виды работ", "Итог за семестр"
                };
                for (int i = 0; i < newPropertyNames.Count; i++)
                {
                    sworksheet.Cell(frow, i + 1).Value = newPropertyNames[i];
                    fworksheet.Cell(frow, i + 1).Value = newPropertyNames[i];
                }

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
