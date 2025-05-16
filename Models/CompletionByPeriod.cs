using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPOVT_DataCenter_Visualisator.Models
{
    public class CompletionByPeriod
    {
        public string? PeriodName { get; set; }
        public decimal CurrentYear { get; set; }
        public decimal PreviousYear { get; set; }
        public string? GrowthPercentage { get; set; }
        public decimal BPCompletion { get; set; }
        public string? BPCompletionPercentage { get; set; }
    }
}
