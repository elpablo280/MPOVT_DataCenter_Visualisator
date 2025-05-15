using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPOVT_DataCenter_Visualisator.Models
{
    public class GeneralInfo
    {
        public GeneralInfo(string name, double planValue, double factValue, string planCompletionPercentage)
        {
            Name = name;
            PlanValue = planValue;
            FactValue = factValue;
            PlanCompletionPercentage = planCompletionPercentage;
        }

        public string Name { get; set; }
        public double PlanValue { get; set; }
        public double FactValue { get; set; }
        public string PlanCompletionPercentage { get; set; }
    }
}
