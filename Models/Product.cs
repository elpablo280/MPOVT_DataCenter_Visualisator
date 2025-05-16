using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPOVT_DataCenter_Visualisator.Models
{
    public class Product
    {
        public string Name { get; set; }
        public int? PlanNumber { get; set; }
        public decimal PlanPrice { get; set; }
        public decimal TotalPlanMoney { get; set; }
        public int? FactNumber { get; set; }
        public decimal TotalFactMoney { get; set; }
        public string PlanCompletionPercentage { get; set; }
        public string Type { get; set; }
        public int? Plan { get; set; }
        public int? Fact { get; set; }
        public int? TotalPlanByToday { get; set; }
        public int? TotalFactByToday { get; set; }
        public int? PlanDeviation { get; set; }
    }
}
