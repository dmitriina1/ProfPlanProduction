using ProfPlan.Commands;
using ProfPlan.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows;

namespace ProfPlan.ViewModels
{
    internal class AddTeacherWindowViewModel
    {
        public string Lastname { get; set; }
        public string Firstname { get; set; }
        public string Middlename { get; set; }
        public string Position { get; set; }
        public string AcademicDegree { get; set; }
        public string Workload { get; set; }
        private bool CanAdd = true;
        private Teacher existingUser { get; set; }

        private RelayCommand _addTeacherCommand;
        public ICommand AddTeacherCommand
        {
            get { return _addTeacherCommand ?? (_addTeacherCommand = new RelayCommand(AddTeacher)); }
        }

        private void AddTeacher(object obj)
        {
            Teacher checkUser = TeachersManager.GetTeacherByName(Lastname, Firstname, Middlename);
            if (existingUser == null && CanAdd == true && checkUser == null)
            {
                double? doubleValue;
                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    if (Workload!=null)
                        if(Workload.IndexOf(",")==-1)
                     doubleValue = Convert.ToDouble(Workload);
                    else
                        {
                            Workload = Workload.Replace(",", ".");
                                doubleValue = Convert.ToDouble(Workload);
                        }
                    else doubleValue = null;
                    TeachersManager.AddTeacher(new Teacher() { LastName = Lastname, FirstName = Firstname, MiddleName = Middlename, Position = Position, AcademicDegree = AcademicDegree, Workload = doubleValue });
                }
                catch (FormatException)
                {
                    MessageBox.Show("В ячейку Ставка было вписано не число!");
                    doubleValue = null;
                }
            }
            else
            {
                if (existingUser == null)
                {
                    existingUser = TeachersManager.GetTeacherByName(Lastname, Firstname, Middlename);

                }
                existingUser.LastName = Lastname;
                existingUser.FirstName = Firstname;
                existingUser.MiddleName = Middlename;
                existingUser.Position = Position;
                existingUser.AcademicDegree = AcademicDegree;
                existingUser.Workload = Workload.ToNullable<double>();

                MessageBox.Show("Данные пользователя обновлены.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }

        }
        public void SetTeacher(Teacher teacher)
        {
            Lastname = teacher?.LastName;
            Firstname = teacher?.FirstName;
            Middlename = teacher?.MiddleName;
            Position = teacher?.Position;
            AcademicDegree = teacher?.AcademicDegree;
            Workload = teacher.Workload.ToNullable<double>().ToString();
            existingUser = TeachersManager.GetTeacherByName(Lastname, Firstname, Middlename);
            if (existingUser == null)
            {
                CanAdd = true;
            }
            else
            {
                CanAdd = false;
            }
        }
    }
}
