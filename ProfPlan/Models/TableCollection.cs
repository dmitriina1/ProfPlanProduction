using ProfPlan.ViewModels;
using ProfPlan.ViewModels.Base;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProfPlan.Models
{
    internal class TableCollection : ViewModel, IEnumerable
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



        public void SubscribeToExcelDataChanges()
        {
            foreach (var excelModel in _excelData)
            {
                excelModel.PropertyChanged -= ExcelModel_PropertyChanged;
            }

            foreach (var excelModel in _excelData)
            {
                excelModel.PropertyChanged += ExcelModel_PropertyChanged;
            }

            UpdateHours();
        }
        public void ExcelModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            UpdateHours();
        }
        private void UpdateHours()
        {
            TotalHours = _excelData.OfType<ExcelModel>().Where(x => x.Total != null).Sum(x => Convert.ToDouble(x.Total));
            AutumnHours = _excelData.OfType<ExcelModel>().Where(x => x.Term != null && x.Term.Equals("нечет", StringComparison.OrdinalIgnoreCase))
                                .Sum(x => Convert.ToDouble(x.Total));
            SpringHours = _excelData.OfType<ExcelModel>().Where(x => x.Term != null && x.Term.Equals("чет", StringComparison.OrdinalIgnoreCase))
                                .Sum(x => Convert.ToDouble(x.Total));
        }
        private double _totalHours;
        public double TotalHours
        {
            get { return _totalHours; }
            set
            {
                if (_totalHours != value)
                {
                    _totalHours = value;
                    OnPropertyChanged(nameof(TotalHours));
                }
            }
        }
        private double _autumnHours;
        public double AutumnHours
        {
            get { return _autumnHours; }
            set
            {
                if (_autumnHours != value)
                {
                    _autumnHours = value;
                    OnPropertyChanged(nameof(AutumnHours));
                }
            }
        }
        private double _springHours;

        public double SpringHours
        {
            get { return _springHours; }
            set
            {
                if (_springHours != value)
                {
                    _springHours = value;
                    OnPropertyChanged(nameof(SpringHours));
                }
            }
        }

        public IEnumerator<ExcelData> GetEnumerator()
        {
            return _excelData.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

    }
}
