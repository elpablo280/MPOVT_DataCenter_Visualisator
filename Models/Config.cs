using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPOVT_DataCenter_Visualisator.Models
{
    public class Config
    {
        public string LogsFilepath { get; set; }
        public string DataCenterFolderpath { get; set; }
        public string ProductionMonitoringFilepath { get; set; }
        public string CronExpression { get; set; }
        public Finreportvalues FinReportValues { get; set; }

        public class Finreportvalues
        {
            public int GeneralInfo_StartRow { get; set; }
            public int GeneralInfo_EndRow { get; set; }
            public int Products_StartRow { get; set; }
            public int Products_EndRow { get; set; }
            public int CompletionByPeriod_StartRow { get; set; }
            public int CompletionByPeriod_EndRow { get; set; }
            public string GeneralInfo_Name_Column { get; set; }
            public string GeneralInfo_PlanValue_Column { get; set; }
            public string GeneralInfo_FactValue_Column { get; set; }
            public string GeneralInfo_PlanCompletionPercentage_Column { get; set; }
            public string Product_Name_Column { get; set; }
            public string Product_PlanNumber_Column { get; set; }
            public string Product_PlanPrice_Column { get; set; }
            public string Product_TotalPlanMoney_Column { get; set; }
            public string Product_FactNumber_Column { get; set; }
            public string Product_TotalFactMoney_Column { get; set; }
            public string Product_PlanCompletionPercentage_Column { get; set; }
            public string Product_Type_Column { get; set; }
            public string Product_Plan_Column { get; set; }
            public string Product_Fact_Column { get; set; }
            public string Product_TotalPlanByToday_Column { get; set; }
            public string Product_TotalFactByToday_Column { get; set; }
            public string Product_PlanDeviation_Column { get; set; }
            public string CompletionByPeriod_PeriodName_Column { get; set; }
            public string CompletionByPeriod_CurrentYear_Column { get; set; }
            public string CompletionByPeriod_PreviousYear_Column { get; set; }
            public string CompletionByPeriod_GrowthPercentage_Column { get; set; }
            public string CompletionByPeriod_BPCompletion_Column { get; set; }
            public string CompletionByPeriod_BPCompletionPercentage_Column { get; set; }
        }
    }
}
