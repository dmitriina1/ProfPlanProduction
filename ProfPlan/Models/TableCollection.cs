using ProfPlan.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlan.Models
{
    internal class TableCollection : ViewModel
    {
        private string _tablename = null;
        private ObservableCollection<ExcelData> _excelData = new ObservableCollection<ExcelData>();

        public ObservableCollection<ExcelData> ExcelData
        {
            get { return _excelData; }
            set
            {
                if (_excelData != value)
                {
                    _excelData = value;
                    OnPropertyChanged(nameof(ExcelData));
                }
            }
        }
        public string Tablename
        {
            get { return _tablename; }
            set
            {
                if (_tablename != value)
                {
                    _tablename = value;
                    OnPropertyChanged(nameof(Tablename));
                }
            }
        }

        private ObservableCollection<TableCollection> _tablesCollection = new ObservableCollection<TableCollection>();
        public ObservableCollection<TableCollection> TablesCollection
        {
            get { return _tablesCollection; }
            set
            {
                if (_tablesCollection != value)
                {
                    _tablesCollection = value;
                    OnPropertyChanged(nameof(TablesCollection));
                }
            }
        }

        public TableCollection(string tablename, ObservableCollection<ExcelData> col)
        {
            Tablename = tablename;
            ExcelData = col;
        }
        public TableCollection(string tablename)
        {
            Tablename = tablename;
        }
        public TableCollection()
        {
        }
    }
}
