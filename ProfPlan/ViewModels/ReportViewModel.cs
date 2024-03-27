using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using ProfPlan.Models;
using ProfPlan.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Diagnostics;
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
        public ObservableCollection<TableCollection> TablesCollectionTeacherSumP { get; set; }
        private ObservableCollection<TableCollection> TablesCollectionTeacherSumListP { get; set; }
        public ObservableCollection<TableCollection> TablesCollectionTeacherSumF { get; set; }
        private ObservableCollection<TableCollection> TablesCollectionTeacherSumListF { get; set; }
        private bool wasCalc = false;
        private int wasCalcIndex = -1;

        //public void CreateLoadCalc(int index)
        //{
        //    SumAllTeachersTables(index);
        //    wasCalc = true;
        //    SaveToExcel(TablesCollectionTeacherSum);
        //}
        public async Task CreateLoadCalcAsync(int index)
        {
            string directoryPath = GetSaveFilePath();
            if (string.IsNullOrEmpty(directoryPath))
                return;
            await Task.Run(() =>
            {
                SumAllTeachersTables(index);
                wasCalc = true;
                wasCalcIndex = index;
                if (index == 0)
                    SaveToExcel(TablesCollectionTeacherSumP, directoryPath);
                else
                    SaveToExcel(TablesCollectionTeacherSumF, directoryPath);

            });
            
        }

        private void SumAllTeachersTables(int index)
        {
            if(index == 0)
            {
                TablesCollectionTeacherSumP = new ObservableCollection<TableCollection>();

                TablesCollectionTeacherSumListP = new ObservableCollection<TableCollection>();
            }
            else
            {
                TablesCollectionTeacherSumF = new ObservableCollection<TableCollection>();

                TablesCollectionTeacherSumListF = new ObservableCollection<TableCollection>();
            }
            

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
                    if (bet == null)
                        bet = totalHours;
                    if (betPercent == 1 || betPercent == null)
                    {
                        sumTableCollection = new TableCollection($"{tableCollection.Tablename}");

                        
                        if (bet!=null)
                        {
                            ProcessBet(index,sumTableCollection,ref sumOddList,ref sumEvenList, bet, betPercent, autumnIndex, springIndex);
                            if(index == 0)
                            TablesCollectionTeacherSumP.Add(sumTableCollection);
                            else
                            TablesCollectionTeacherSumF.Add(sumTableCollection);
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

                                ProcessBet(index,sumTableCollection,ref sumOddList,ref sumEvenList, bet, betPercent, autumnIndex, springIndex);
                                if (index == 0)
                                    TablesCollectionTeacherSumP.Add(sumTableCollection);
                                else
                                    TablesCollectionTeacherSumF.Add(sumTableCollection);

                                ProcessBet(index,sumTableCollectionTwo,ref sumOddList,ref sumEvenList, bet*(betPercent - 1), betPercent, autumnIndex, springIndex);
                                if (index == 0)
                                    TablesCollectionTeacherSumP.Add(sumTableCollectionTwo);
                                else
                                    TablesCollectionTeacherSumF.Add(sumTableCollectionTwo);
                            }
                            else
                            {
                                    sumTableCollection = new TableCollection($"{tableCollection.Tablename} {betPercent}");

                                    ProcessBet(index,sumTableCollection,ref sumOddList,ref sumEvenList, bet, betPercent, autumnIndex, springIndex);
                                    if (index == 0)
                                        TablesCollectionTeacherSumP.Add(sumTableCollection);
                                    else
                                        TablesCollectionTeacherSumF.Add(sumTableCollection);

                            }
                        }
                        
                    }                    


                }

            }
            
        }


        private TableCollection ProcessBet(int index, TableCollection sumTableCollection,ref ObservableCollection<ExcelModel> sumOddList,ref ObservableCollection<ExcelModel> sumEvenList, double? bet, double? betPercent, double? autumnIndex = null, double? springIndex = null)
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
            var sortedList = sumOddList.OrderByDescending(x => x.SumProperties());

            foreach (ExcelModel excelModel in sortedList)
            {
                if (betValue > sum)
                {
                    dif = betValue - sum;
                    // Если разница между betValue и текущей суммой больше или равна свойству SumProperties
                    if (dif >= excelModel.SumProperties())
                    {
                        sum += excelModel.SumProperties();
                        sumOddListOneBet.ExcelData.Add(excelModel);
                        ListForIPPlan.ExcelData.Add(excelModel);
                    }
                }
                else
                {
                    break;
                }
            }
            DeleteItemsFromObsCol(ref sumOddList, sumOddListOneBet);
            
            sum = 0;
            min = null;
            betValue = bet * springIndex;
            if (betPercent<1)
            {
                betValue = bet * springIndex * betPercent;
            }

            sortedList = sumEvenList.OrderByDescending(x => x.SumProperties());

            foreach (ExcelModel excelModel in sortedList)
            {
                if (betValue > sum)
                {
                    dif = betValue - sum;
                    // Если разница между betValue и текущей суммой больше или равна свойству SumProperties
                    if (dif >= excelModel.SumProperties())
                    {
                        sum += excelModel.SumProperties();
                        sumEvenListOneBet.ExcelData.Add(excelModel);
                        ListForIPPlan.ExcelData.Add(excelModel);
                    }
                }
                else
                {
                    break;
                }
            }
            DeleteItemsFromObsCol(ref sumEvenList, sumEvenListOneBet);

            //List<int> lists = new List<int>();
            //for (int i = 0; i<sumOddListOneBet.ExcelData.Count; i++)
            //{
            //    lists.Add((sumOddListOneBet.ExcelData[i] as ExcelModel).Number);
            //}
            //MessageBox.Show($"{sumTableCollection.Tablename}{string.Join(", ", lists)}");
            //lists.Clear();
            //for (int i = 0; i<sumEvenListOneBet.ExcelData.Count; i++)
            //{
            //    lists.Add((sumEvenListOneBet.ExcelData[i] as ExcelModel).Number);
            //}
            //MessageBox.Show($"{sumTableCollection.Tablename}{string.Join(", ", lists)}");

            ExcelModel sumOddOneBet = CalculateSum(sumOddListOneBet, "нечет");
            ExcelModel sumEvenOneBet = CalculateSum(sumEvenListOneBet, "чет");
            sumTableCollection.ExcelData.Add(sumOddOneBet);
            sumTableCollection.ExcelData.Add(sumEvenOneBet);
            if(index == 0)
            TablesCollectionTeacherSumListP.Add(ListForIPPlan);
            else
                TablesCollectionTeacherSumListF.Add(ListForIPPlan);
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


        private void DeleteItemsFromObsCol(ref ObservableCollection<ExcelModel> collection, TableCollection tabCol)
        {
            foreach(ExcelModel tab in tabCol)
            {
                collection.Remove(tab);
            }
        }


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

        public void SaveToExcel(ObservableCollection<TableCollection> tablesCollection, string directoryPath)
        {
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

                        fworksheet.Cell(rowNumber, columnNumber).Value = totalSemester.ToString();

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


                        fworksheet.Cell(rowNumber, cnum).Value = totalSemester.ToString();

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
                fworksheet.Cell(row, 42).Value = (value21 + value41).ToString();
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
        private List<IndividualPlan> CreateIPList(TableCollection tab, IXLWorkbook workbook)
        {
            List<IndividualPlan> IPList = new List<IndividualPlan>();
            //Перечень предметов

            foreach (ExcelModel excel in tab.ExcelData)
            {
                if (excel.Total != 0 && excel.Total!=null)
                {
                    IPList.Add(excel.FormulateIndividualPlan());
                    IPList[IPList.Count - 1].TypeOfWork = excel.GetTypeOfWork();
                }
            }
            var groupedPlans = IPList.GroupBy(ip => new { ip.Discipline, ip.TypeOfWork, ip.Term, ip.Group, ip.GroupCount, ip.SubGroup, ip.Branch })
                                .Select(group => new IndividualPlan(
                                    group.Key.Discipline,
                                    group.Key.TypeOfWork,
                                    group.Key.Term,
                                    group.Key.Group,
                                    group.Key.GroupCount,
                                    group.Key.SubGroup,
                                    group.Key.Branch,
                                    group.Sum(ip => ip.Hours)
                                ))
                                .ToList();

            IPList = groupedPlans;
            IPList = IPList.OrderBy(ip => ip.Discipline)
           .ThenBy(ip => ip.TypeOfWork)
           .ThenBy(ip => ip.SubGroup)
           .ThenBy(ip => ip.Group)

           .ToList();
            return IPList;
        }
        public void WorkWithWorkSheet(IXLWorksheet worksheet, int row, List<IndividualPlan> IPList )
        {
            

            //
            worksheet.Cell(row, 1).Value = "Четный семестр";
            row++;
            worksheet.Cell(row, 1).Value = "Дисциплина";
            worksheet.Cell(row, 2).Value = "Вид работы";
            worksheet.Cell(row, 3).Value = "Группа";
            worksheet.Cell(row, 4).Value = "Подгруппа";
            worksheet.Cell(row, 5).Value = "Филиал";
            worksheet.Cell(row, 6).Value = "Часы";
            row++;
            foreach (IndividualPlan ip in IPList)
            {
                if (ip.Term.IndexOf("нечет", StringComparison.OrdinalIgnoreCase) == -1)
                {
                    worksheet.Cell(row, 1).Value = ip.Discipline;
                    worksheet.Cell(row, 2).Value = ip.TypeOfWork;
                    worksheet.Cell(row, 3).Value = ip.Group;
                    worksheet.Cell(row, 4).Value = ip.SubGroup;
                    worksheet.Cell(row, 5).Value = ip.Branch;
                    worksheet.Cell(row, 6).Value = ip.Hours;
                    row++;
                }
            }
            row+=2;

            worksheet.Cell(row, 1).Value = "Нетный семестр";
            row++;
            worksheet.Cell(row, 1).Value = "Дисциплина";
            worksheet.Cell(row, 2).Value = "Вид работы";
            worksheet.Cell(row, 3).Value = "Группа";
            worksheet.Cell(row, 4).Value = "Подгруппа";
            worksheet.Cell(row, 5).Value = "Филиал";
            worksheet.Cell(row, 6).Value = "Часы";
            row++;
            foreach (IndividualPlan ip in IPList)
            {
                if (ip.Term.IndexOf("нечет", StringComparison.OrdinalIgnoreCase) != -1)
                {
                    worksheet.Cell(row, 1).Value = ip.Discipline;
                    worksheet.Cell(row, 2).Value = ip.TypeOfWork;
                    worksheet.Cell(row, 3).Value = ip.Group;
                    worksheet.Cell(row, 4).Value = ip.SubGroup;
                    worksheet.Cell(row, 5).Value = ip.Branch;
                    worksheet.Cell(row, 6).Value = ip.Hours;
                    row++;
                }
            }
            worksheet.Columns().AdjustToContents();
            worksheet.Rows().AdjustToContents();
        }
        private ObservableCollection<List<IndividualPlan>> IPListP = new ObservableCollection<List<IndividualPlan>>(), IPListF = new ObservableCollection<List<IndividualPlan>>();
        public async Task CreateIndividualPlan(int index)
        {
            var workbook = new XLWorkbook();

            string directoryPath = GetSaveFilePathForIP();
            if (string.IsNullOrEmpty(directoryPath))
                return;
            await Task.Run(() =>
            {
                if (wasCalc == false && wasCalcIndex != index)
                {
                    SumAllTeachersTables(index);
                }
                if(index == 0)
                {
                    SumAllTeachersTables(1);
                }
                else
                {
                    SumAllTeachersTables(0);
                }
                if(TablesCollectionTeacherSumListP!= null && TablesCollectionTeacherSumListP.Count>0)
                foreach (TableCollection tab in TablesCollectionTeacherSumListP)
                {
                        IPListP.Add(CreateIPList(tab, workbook));
                    }
                if (TablesCollectionTeacherSumListF!=null && TablesCollectionTeacherSumListF.Count>0)
                    foreach (TableCollection tab in TablesCollectionTeacherSumListF)
                {
                        IPListF.Add(CreateIPList(tab, workbook));
                    }
                ObservableCollection<TableCollection> SomeTab;
                if (IPListF != null && IPListF.Count>0)
                    SomeTab = TablesCollectionTeacherSumListF;
                else
                    SomeTab = TablesCollectionTeacherSumListP;
                for (int i = 0;i< SomeTab.Count;i++)
                {
                    int row = 1;
                    string tabName = SomeTab[i].Tablename;
                    if (SomeTab[i].Tablename.StartsWith("П_") || SomeTab[i].Tablename.StartsWith("Ф_"))
                    {
                        tabName = SomeTab[i].Tablename.Substring(2); // Удаление "П_" или "Ф_" из начала строки
                    }
                    var worksheet = workbook.Worksheets.Add(tabName); 

                    //Итого
                    worksheet.Range(row, 1, row + 1, 1).Merge();
                    worksheet.Cell(row, 1).Value = "Виды учебных занятий (работ)";
                   
                    worksheet.Range(row, 2, row, 3).Merge();
                    worksheet.Cell(row, 2).Value = "нечетный семестр";

                    worksheet.Range(row, 4, row, 5).Merge();
                    worksheet.Cell(row, 4).Value = "четный семестр";

                    worksheet.Range(row, 6, row, 7).Merge();
                    worksheet.Cell(row, 6).Value = "Итого за уч.год";
                    row++;
                    worksheet.Cell(row, 2).Value = "План";
                    worksheet.Cell(row, 4).Value = "План";
                    worksheet.Cell(row, 6).Value = "План";
                    worksheet.Cell(row, 7).Value = "Факт";
                    worksheet.Cell(row, 3).Value = "Факт";
                    worksheet.Cell(row, 5).Value = "Факт";

                    row++;
                    int r = row;

                    int col1, col2, col3;
                    if (index == 0)
                    {
                        col1 = 2;
                        col2 = 4;
                        col3 = 6;
                    }
                    else
                    {
                        col1 = 3;
                        col2 = 5;
                        col3 = 7;
                    }
                    //var groupedByTypeOfWork = IPList.GroupBy(ip => new { ip.TypeOfWork, ip.Term });
                    //List<(string TypeOfWork, string Term, double? TotalHours)> resultList = new List<(string, string, double?)>();
                    if (IPListF.Count>0 && IPListP.Count>0)
                    {
                        col1 = 2;
                        col2 = 4;
                        col3 = 6;
                        CreateTotal(IPListP[i], index, r, ref row, worksheet, col1, col2, col3);
                        col1 = 3;
                        col2 = 5;
                        col3 = 7;
                        CreateTotal(IPListF[i], index, r, ref row, worksheet, col1, col2, col3);
                        SumAllTables(worksheet, row, r);
                        row+=2;

                        WorkWithWorkSheet(worksheet, row, IPListF[i]);
                    }
                    else if (IPListF.Count>0)
                    {
                        CreateTotal(IPListF[i], index, r, ref row, worksheet, col1, col2, col3);
                        SumTables(worksheet, row, r, col1, col2, col3);
                        row+=2;

                        WorkWithWorkSheet(worksheet, row, IPListF[i]);
                    }
                    else
                    {
                        CreateTotal(IPListP[i], index, r, ref row, worksheet, col1, col2, col3);
                        SumTables(worksheet, row, r, col1, col2, col3);
                        row+=2;

                        WorkWithWorkSheet(worksheet, row, IPListP[i]);
                    }
                    

                     
                }
            });

            
            workbook.SaveAs(directoryPath);
        }
        private void CreateTotal(List<IndividualPlan> IPList, int index, int r, ref int row, IXLWorksheet worksheet, int col1, int col2, int col3)
        {
            int ind;
            var evenTermList = IPList.Where(ip => ip.Term == "чет").ToList();
            var oddTermList = IPList.Where(ip => ip.Term == "нечет").ToList();

            // Группировка и подсчет суммы часов для каждого типа работы
            var evenTermGrouped = evenTermList.GroupBy(ip => ip.TypeOfWork)
                                              .Select(group => new { TypeOfWork = group.Key, TotalHours = group.Sum(ip => ip.Hours) })
                                              .ToList();
            var oddTermGrouped = oddTermList.GroupBy(ip => ip.TypeOfWork)
                                            .Select(group => new { TypeOfWork = group.Key, TotalHours = group.Sum(ip => ip.Hours) })
                                            .ToList();
            
            
            foreach (var group in oddTermGrouped)
            {
                var existingRow = worksheet.RowsUsed().FirstOrDefault(s => s.Cell(1).Value.ToString() == group.TypeOfWork);
                if (existingRow == null)
                {
                    worksheet.Cell(row, 1).Value = group.TypeOfWork;
                    worksheet.Cell(row, col1).Value = group.TotalHours;
                    row++;
                }
                else
                {
                    existingRow.Cell(col1).Value = group.TotalHours;
                }
            }
            foreach (var group in evenTermGrouped)
            {
                var existingRow = worksheet.RowsUsed().FirstOrDefault(s => s.Cell(1).Value.ToString() == group.TypeOfWork);
                if (existingRow == null)
                {
                    worksheet.Cell(row, 1).Value = group.TypeOfWork;
                    worksheet.Cell(row, col2).Value = group.TotalHours;
                    row++;
                }
                else
                {
                    existingRow.Cell(col2).Value = group.TotalHours;
                }

            }
        }
        private void SumTables(IXLWorksheet worksheet, int row, int r, int col1, int col2, int col3)
        {
            double? sum;
            for (int i = r; i < row; i++)
            {
                var range1 = string.IsNullOrEmpty(worksheet.Row(i).Cell(col1).Value.ToString()) ? 0 : Convert.ToDouble(worksheet.Row(i).Cell(col1).Value.ToString());
                var range2 = string.IsNullOrEmpty(worksheet.Row(i).Cell(col2).Value.ToString()) ? 0 : Convert.ToDouble(worksheet.Row(i).Cell(col2).Value.ToString());
                sum = Convert.ToDouble(range1) + Convert.ToDouble(range2);
                worksheet.Cell(i, col3).Value = sum;
            }
            worksheet.Cell(row, 1).Value = "Итого";
            //MessageBox.Show(worksheet.Cell(3, col1).Value.ToString());
            //MessageBox.Show(worksheet.Cell(row, col1).Value.ToString());
            var range = worksheet.Range(3, col1, row, col1);
            sum = range.CellsUsed().Sum(cell => cell.GetDouble());
            worksheet.Cell(row, col1).Value = sum;
            range = worksheet.Range(3, col2, row, col2);
            sum = range.CellsUsed().Sum(cell => cell.GetDouble());
            worksheet.Cell(row, col2).Value = sum;
            range = worksheet.Range(3, col3, row, col3);
            sum = range.CellsUsed().Sum(cell => cell.GetDouble());
            worksheet.Cell(row, col3).Value = sum;
        }

        private void SumAllTables(IXLWorksheet worksheet, int row, int r)
        {
            SumTables(worksheet, row, r, 2, 4, 6);
            SumTables(worksheet, row, r, 3, 5, 7);
        }

        private string GetSaveFilePathForIP()
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.Title = "Save Excel File";
            saveFileDialog.FileName = $"ИП {DateTime.Today:dd-MM-yyyy}.xlsx";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                return saveFileDialog.FileName;

            return null;
        }

    }
}