using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProfPlan.Models
{
    internal class IndividualPlan
    {
        public string Discipline { get; set; }
        public string TypeOfWork { get; set; }
        public string Term { get; set; }
        public string Group { get; set; }
        public string Branch { get; set; }
        public int? GroupCount { get; set; }
        public string SubGroup { get; set; }
        public double? Hours { get; set; }
        public IndividualPlan(string discipline, string typeofwork, string term, string group, int? groupCount, string subGroup, string branch, double? hours)
        {
            Discipline=discipline;
            TypeOfWork = typeofwork;
            Term=term;
            Group=group;
            Branch=branch;
            GroupCount=groupCount;
            SubGroup=subGroup;
            Hours=hours;
        }
    }
}
