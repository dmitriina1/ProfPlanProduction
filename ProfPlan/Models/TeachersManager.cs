using DocumentFormat.OpenXml.Wordprocessing;
using ProfPlan.ViewModels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ProfPlan.Models
{
    internal class TeachersManager
    {
        public static ObservableCollection<Teacher> _DatabaseUsers = new ObservableCollection<Teacher>();

        public static ObservableCollection<Teacher> GetTeachers()
        {
            var teacherDatabase = new TeacherDatabase();
            _DatabaseUsers = teacherDatabase.LoadTeachers();
            return _DatabaseUsers;

        }
        public TeachersManager()
        {
            var teacherDatabase = new TeacherDatabase();
            _DatabaseUsers = teacherDatabase.LoadTeachers();
        }

        public static void AddTeacher(Teacher teacher)
        {
            _DatabaseUsers.Add(teacher);
            var teacherDatabase = new TeacherDatabase();
            teacherDatabase.SaveTeachers(_DatabaseUsers);
            ExcelModel.UpdateSharedTeachers();
            OnTeachersChanged();
        }
        public static Teacher GetTeacherByName(string lastname, string firstname, string middlename)
        {
            return _DatabaseUsers.FirstOrDefault(teacher => teacher.LastName == lastname && teacher.FirstName == firstname && teacher.MiddleName == middlename);
        }
        public static int GetTeacherIndex(Teacher teach)
        {
            return _DatabaseUsers.IndexOf(_DatabaseUsers.FirstOrDefault(teacher => teacher.LastName == teach.LastName && teacher.FirstName == teach.FirstName && teacher.MiddleName == teach.MiddleName));
        }
        public static void UpdateTeacher(Teacher teacher, int index)
        {
            _DatabaseUsers[index] = teacher;
            var teacherDatabase = new TeacherDatabase();
            teacherDatabase.SaveTeachers(_DatabaseUsers);
            ExcelModel.UpdateSharedTeachers();
            OnTeachersChanged();
        }
        public static void RemoveTeacher(Teacher teacher)
        {
            _DatabaseUsers.Remove(teacher);
            var teacherDatabase = new TeacherDatabase();
            teacherDatabase.SaveTeachers(_DatabaseUsers);
            ExcelModel.UpdateSharedTeachers();
            OnTeachersChanged();
        }

        public static event EventHandler TeachersChanged;

        private static void OnTeachersChanged()
        {
            TeachersChanged?.Invoke(null, EventArgs.Empty);
        }
    }
}
