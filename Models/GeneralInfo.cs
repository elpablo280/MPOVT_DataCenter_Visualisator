using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPOVT_DataCenter_Visualisator.Models
{
    public class GeneralInfo
    {
        public string? Name { get; set; }
        public decimal PlanValue { get; set; }
        public decimal FactValue { get; set; }
        public string? PlanCompletionPercentage { get; set; }
    }
}
