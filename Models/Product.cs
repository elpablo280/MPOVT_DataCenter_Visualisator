using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPOVT_DataCenter_Visualisator.Models
{
    public class Product(
        string name,
        int? planNumber,
        decimal planPrice,
        decimal totalPlanMoney,
        int? factNumber,
        decimal totalFactMoney,
        string planCompletionPercentage,
        string type,
        int? plan,
        int? fact,
        int? totalPlanByToday,
        int? totalFactByToday,
        int? planDeviation
            )
    {
        public string Name { get; set; } = name;
        public int? PlanNumber { get; set; } = planNumber;
        public decimal PlanPrice { get; set; } = planPrice;
        public decimal TotalPlanMoney { get; set; } = totalPlanMoney;
        public int? FactNumber { get; set; } = factNumber;
        public decimal TotalFactMoney { get; set; } = totalFactMoney;
        public string PlanCompletionPercentage { get; set; } = planCompletionPercentage;
        public string Type { get; set; } = type;
        public int? Plan { get; set; } = plan;
        public int? Fact { get; set; } = fact;
        public int? TotalPlanByToday { get; set; } = totalPlanByToday;
        public int? TotalFactByToday { get; set; } = totalFactByToday;
        public int? PlanDeviation { get; set; } = planDeviation;
    }
}
