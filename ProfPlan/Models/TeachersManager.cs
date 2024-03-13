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
            return _DatabaseUsers;

        }


        public static void AddTeacher(Teacher teacher)
        {
            _DatabaseUsers.Add(teacher);

        }
        public static Teacher GetTeacherByName(string lastname, string firstname, string middlename)
        {
            return _DatabaseUsers.FirstOrDefault(teacher => teacher.LastName == lastname && teacher.FirstName == firstname && teacher.MiddleName == middlename);
        }
    }
}
