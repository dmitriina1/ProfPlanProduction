using ProfPlan.Commands;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows;
using ProfPlan.Models;
using ProfPlan.Views;
using ProfPlan.ViewModels.Base;

namespace ProfPlan.ViewModels
{
    internal class TeachersWindowViewModel : ViewModel
    {
        public ObservableCollection<Teacher> Teachers { get; set; }

        public ICommand ShowWindowCommand { get; set; }



        public TeachersWindowViewModel()
        {
            Teachers = TeachersManager.GetTeachers();

            ShowWindowCommand = new RelayCommand(ShowWindow, CanShowWindow);

        }

        private bool CanShowWindow(object obj)
        {
            return true;
        }

        private void ShowWindow(object obj)
        {
            var mainWindow = obj as Window;

            AddTeacherWindow addTeacherWin = new AddTeacherWindow();
            addTeacherWin.Owner = mainWindow;
            addTeacherWin.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            addTeacherWin.ShowDialog();


        }


        private Teacher _selectedTeacher;
        public Teacher SelectedTeacher
        {
            get { return _selectedTeacher; }
            set
            {
                _selectedTeacher = value;
                OnPropertyChanged(nameof(SelectedTeacher));
            }
        }



        public void RemoveSelectedTeacher(Teacher teacher)
        {
            if (MessageBox.Show($"Вы уверены, что хотите удалить пользователя {teacher.LastName} {teacher.FirstName} {teacher.MiddleName}?", "Удаление пользователя", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                // Удаление пользователя из коллекции и обновление представления
                Teachers.Remove(teacher);
            }
        }
    }
}
