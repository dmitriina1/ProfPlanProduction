using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        }
        public static Teacher GetTeacherByName(string lastname, string firstname, string middlename)
        {
            return _DatabaseUsers.FirstOrDefault(teacher => teacher.LastName == lastname && teacher.FirstName == firstname && teacher.MiddleName == middlename);
        }
    }
}
