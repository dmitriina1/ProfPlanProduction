using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlan.Models
{
    internal class ExcelTotal : ExcelData
    {
        public string Teacher { get; set; }
        private int? _bet;
        public int? Bet
        {
            get { return _bet; }
            set
            {
                if (_bet != value)
                {
                    _bet = value;
                    OnPropertyChanged(nameof(Bet));
                    if (Bet != null && TotalHours != null)
                    {
                        Difference = Math.Round(Convert.ToDouble(TotalHours - Bet), 2);
                    }
                    else
                    {
                        Difference = 0;
                    }
                }
            }
        }
        public double? BetPercent { get; set; }
        private double? _totalHours;
        public double? TotalHours
        {
            get { return _totalHours; }
            set
            {
                if (_totalHours != value)
                {
                    _totalHours = value;
                    OnPropertyChanged(nameof(TotalHours));
                    if (Bet != null && TotalHours != null)
                    {
                        Difference = Math.Round(Convert.ToDouble(TotalHours - Bet), 2);
                    }
                    else
                    {
                        Difference = 0;
                    }
                }
            }
        }
        public double? AutumnHours { get; set; }
        public double? SpringHours { get; set; }
        public ExcelTotal() { }
        public ExcelTotal(string techer, int? bet, double? betpercent, double? total, double? autumnhours, double? springHours, double? difference)
        {
            Teacher = techer;
            Bet = bet;
            BetPercent = betpercent;
            TotalHours = total;
            AutumnHours = autumnhours;
            SpringHours = springHours;
            Difference = difference;
        }

        private double? _difference;
        public double? Difference
        {
            get { return _difference; }
            set
            {
                if (_difference != value)
                {
                    if (value != 0)
                    {
                        _difference = value;
                    }
                    else _difference = null;
                    OnPropertyChanged(nameof(Difference));
                }
            }
        }
    }
}

