    using ProfPlan.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlan.ViewModels
{
    internal class ExcelModel : ExcelData
    {
        private static ObservableCollection<string> sharedTeachers = new ObservableCollection<string>();

        public static ObservableCollection<string> Teachers
        {
            get { return sharedTeachers; }
            set
            {
                if (sharedTeachers != value)
                {
                    sharedTeachers = value;
                }
            }
        }
        public static void AddToSharedTeachers(string teacher)
        {
            sharedTeachers.Add(teacher);
        }
        public static void UpdateSharedTeachers()
        {
            sharedTeachers.Clear();
            string lname, fname, mname;
            foreach(var teacher in TeachersManager.GetTeachers())
            {
                lname=teacher.LastName;
                fname=teacher.FirstName;
                mname=teacher.MiddleName;
                if(mname.Length > 0) 
                    sharedTeachers.Add($"{lname} {fname[0]}.{mname[0]}.");
                else
                    sharedTeachers.Add($"{lname} {fname[0]}.");
            }
        }
        public int Number { get; set; }
        private string _teacher;

        public string Teacher
        {
            get { return _teacher; }
            set
            {
                if (_teacher != value)
                {
                    _teacher = value;
                    OnPropertyChanged(nameof(Teacher));
                }
            }
        }
        public string Discipline { get; set; }
        public string Term { get; set; }
        public string Group { get; set; }
        public string Institute { get; set; }
        public int? GroupCount { get; set; }
        public string SubGroup { get; set; }
        public string FormOfStudy { get; set; }
        public int? StudentsCount { get; set; }
        public int? CommercicalStudentsCount { get; set; }
        public int? Weeks { get; set; }
        public string ReportingForm { get; set; }
        public double? Lectures { get; set; }
        public double? Practices { get; set; }
        public double? Laboratory { get; set; }
        public double? Consultations { get; set; }
        public double? Tests { get; set; }
        public double? Exams { get; set; }
        public double? CourseWorks { get; set; }
        public double? CourseProjects { get; set; }
        public double? GEKAndGAK { get; set; }
        public double? Diploma { get; set; }
        public double? RGZ { get; set; }
        public double? ReviewDiploma { get; set; }
        public double? Other { get; set; }
        public double? Total { get; set; }
        public double? Budget { get; set; }
        public double? Commercial { get; set; }
        public ExcelModel(
            int number, string teacher, string discipline, string term,
            string group, string institute, int? groupCount, string subGroup,
            string formOfStudy, int? studentsCount, int? commercicalStudentsCount,
            int? weeks, string reportingForm, double? lectures, double? practices,
            double? laboratory, double? consultations, double? tests, double? exams,
            double? courseWorks, double? courseProjects, double? gEKAndGAK, double? diploma,
            double? rGZ, double? reviewDiploma, double? other, double? total, double? budget,
            double? commercial)
        {
            Number = number;
            Teacher = teacher;
            Discipline = discipline;
            Term = term;
            Group = group;
            Institute = institute;
            GroupCount = groupCount;
            SubGroup = subGroup;
            FormOfStudy = formOfStudy;
            StudentsCount = studentsCount;
            CommercicalStudentsCount = commercicalStudentsCount;
            Weeks = weeks;
            ReportingForm = reportingForm;
            Lectures = lectures;
            Practices = practices;
            Laboratory = laboratory;
            Consultations = consultations;
            Tests = tests;
            Exams = exams;
            CourseWorks = courseWorks;
            CourseProjects = courseProjects;
            GEKAndGAK = gEKAndGAK;
            Diploma = diploma;
            RGZ = rGZ;
            ReviewDiploma = reviewDiploma;
            Other = other;
            Total = total;
            Budget = budget;
            Commercial = commercial;
        }

        public double SumProperties()
        {
            return (Lectures ?? 0) +
                       (Consultations ?? 0) +
                       (Laboratory ?? 0) +
                       (Practices ?? 0) +
                       (Tests ?? 0) +
                       (Exams ?? 0) +
                       (CourseProjects ?? 0) +
                       (CourseWorks ?? 0) +
                       (Diploma ?? 0) +
                       (RGZ ?? 0) +
                       (GEKAndGAK ?? 0) +
                       (ReviewDiploma ?? 0) +
                       (Other ?? 0);
        }

        public string GetTypeOfWork()
        {
            var properties = new List<(string Name, double? Value)>
    {
        ("Чтение лекций", Lectures),
        ("Проведение практических занятий", Practices),
        ("Проведение лабораторных работ", Laboratory),
        ("Проведение консультаций", Consultations),
        ("Прием зачетов переаттестаций", Tests),
        ("Экзамен семестровый", Exams),
        ("Курсовые работы", CourseWorks),
        ("Курсовые проекты", CourseProjects),
        ("Проведение ГЭК", GEKAndGAK),
        ("ВКР бакалавров", Diploma),
        ("Прием РГЗ", RGZ),
        ("Практическая работа/Рецензирование диплома", ReviewDiploma),
        ("Прочее", Other)
    };

            var nonZeroProperty = properties.FirstOrDefault(p => p.Value != null);

            if (nonZeroProperty != default)
            {
                if (nonZeroProperty.Name == "Практическая работа/Рецензирование диплома")
                {
                    // Проверяем, содержит ли строка Discipline подстроку "преддипломная практика"
                    if (Discipline != null && Discipline.Contains("преддипломная практика"))
                    {
                        return "Рецензирование диплома";
                    }
                    else
                    {
                        return "Практическая работа";
                    }
                }
                else
                {
                    return nonZeroProperty.Name;
                }
            }
            else
            {
                return null;
            }
        }

        public IndividualPlan FormulateIndividualPlan()
        {
            return new IndividualPlan(Discipline, GetTypeOfWork(), Term, Group, GroupCount, SubGroup, $"СГУГиТ ({Institute})", Total);
        }
    }
}

