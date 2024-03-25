using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Data.SqlClient;
using System.IO;
using System.Data;

namespace ProfPlan.Models
{
    internal class TeacherDatabase
    {
        private string connectionString;

        public TeacherDatabase()
        {
            InitializeDatabase();
        }

        private String dbFileName;
        private SQLiteConnection m_dbConn;
        private SQLiteCommand m_sqlCmd;
        // Метод для инициализации базы данных и создания таблицы, если она не существует
        private void InitializeDatabase()
        {
            m_dbConn = new SQLiteConnection();
            m_sqlCmd = new SQLiteCommand();
            dbFileName = "Teachers.db";
            if (!File.Exists(dbFileName))
            {
                SQLiteConnection.CreateFile(dbFileName);
            }
            m_dbConn = new SQLiteConnection("Data Source=" + dbFileName + ";Version=3;");
            m_dbConn.Open();
            m_sqlCmd.Connection = m_dbConn;

            m_sqlCmd.CommandText = @"
                    CREATE TABLE IF NOT EXISTS Teachers (
                        LastName TEXT,
                        FirstName TEXT,
                        MiddleName TEXT,
                        Position TEXT,
                        AcademicDegree TEXT,
                        Workload REAL
                    );"; 
            m_sqlCmd.ExecuteNonQuery();
           
        }

        // Метод для сохранения данных учителя в базу данных
        public void SaveTeachers(ObservableCollection<Teacher> teachers)
        {

            if (m_dbConn.State != ConnectionState.Open)
            {
                MessageBox.Show("Open connection with database");
                return;
            }
            m_sqlCmd.CommandText = "DELETE FROM Teachers";
            m_sqlCmd.ExecuteNonQuery();
            foreach (var teacher in teachers)
            {

                m_sqlCmd.CommandText = @"
                        INSERT INTO Teachers (LastName, FirstName, MiddleName, Position, AcademicDegree, Workload)
                        VALUES ($lastName, $firstName, $middleName, $position, $academicDegree, $workload);";

                m_sqlCmd.Parameters.AddWithValue("$lastName", teacher.LastName);
                m_sqlCmd.Parameters.AddWithValue("$firstName", teacher.FirstName);
                m_sqlCmd.Parameters.AddWithValue("$middleName", teacher.MiddleName);
                m_sqlCmd.Parameters.AddWithValue("$position", teacher.Position);
                m_sqlCmd.Parameters.AddWithValue("$academicDegree", teacher.AcademicDegree);
                m_sqlCmd.Parameters.AddWithValue("$workload", teacher.Workload);

                m_sqlCmd.ExecuteNonQuery();
            }
        }

        public ObservableCollection<Teacher> LoadTeachers()
        {
            ObservableCollection<Teacher> teachers = new ObservableCollection<Teacher>();

            try
            {
                m_sqlCmd.CommandText = "SELECT * FROM Teachers";
                SQLiteDataReader reader = m_sqlCmd.ExecuteReader();

                while (reader.Read())
                {
                    string lastName = reader["LastName"].ToString();
                    string firstName = reader["FirstName"].ToString();
                    string middleName = reader["MiddleName"].ToString();
                    string position = reader["Position"].ToString();
                    string academicDegree = reader["AcademicDegree"].ToString();
                    double? workload = reader["Workload"].ToNullable<double>();

                    Teacher teacher = new Teacher(lastName, firstName, middleName, position, academicDegree, workload);
                    teachers.Add(teacher);
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading teachers from database: " + ex.Message);
            }

            return teachers;
        }

    }
}
