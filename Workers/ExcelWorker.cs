using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using MPOVT_DataCenter_Visualisator.Jobs;
using MPOVT_DataCenter_Visualisator.Models;

namespace MPOVT_DataCenter_Visualisator.Workers
{
    public class ExcelWorker
    {
        private XLWorkbook? Wb;
        private readonly string Filepath;
        private readonly Config Config;

        public ExcelWorker(string filepath, Config config)
        {
            Filepath = filepath;
            Config = config;
        }

        private void OpenWorkbook()
        {
            Wb = new(Filepath);
        }

        public (List<GeneralInfo>, List<Product>, List<CompletionByPeriod>) ImportData()
        {
            OpenWorkbook();
            List<GeneralInfo> GeneralInfoList = new();
            List<Product> Products = new();
            List<CompletionByPeriod> CompletionByPeriodList = new();
            using (Wb)
            {
                var ws = Wb.Worksheet(3);   // есть скрытые листы

                //ws.CollapseRows();
                //ws.CollapseColumns();

                // читаем данные
                GeneralInfoList = GetTable<GeneralInfo>(ws, Config.FinReportValues.GeneralInfo_StartRow, Config.FinReportValues.GeneralInfo_EndRow);
                Products = GetTable<Product>(ws, Config.FinReportValues.Products_StartRow, Config.FinReportValues.Products_EndRow);
                CompletionByPeriodList = GetTable<CompletionByPeriod>(ws, Config.FinReportValues.CompletionByPeriod_StartRow, Config.FinReportValues.CompletionByPeriod_EndRow);
            }

            return (GeneralInfoList, Products, CompletionByPeriodList);
        }

        private List<T> GetTable<T>(IXLWorksheet ws, int startRow, int endRow) where T : class, new()
        {
            List<T> resultList = new();
            switch (typeof(T).Name)
            {
                case nameof(GeneralInfo):
                    List<GeneralInfo> GeneralInfoList = new();
                    foreach (var row in ws.Rows(startRow, endRow))
                    {
                        if (!row.IsHidden)
                        {
                            T item = new T();

                            row.Cell(Config.FinReportValues.GeneralInfo_Name_Column).TryGetValue(out string name);
                            row.Cell(Config.FinReportValues.GeneralInfo_PlanValue_Column).TryGetValue(out decimal planValue);
                            row.Cell(Config.FinReportValues.GeneralInfo_FactValue_Column).TryGetValue(out decimal factValue);
                            string? planCompletionPercentage = GetPlanCompletionPercentageString(row.Cell(Config.FinReportValues.GeneralInfo_PlanCompletionPercentage_Column));

                            GeneralInfo generalInfo = new()
                            {
                                Name = name,
                                PlanValue = Math.Round(planValue, 2),
                                FactValue = Math.Round(factValue, 2),
                                PlanCompletionPercentage = planCompletionPercentage,
                            };
                            resultList.Add(generalInfo as T);
                        }
                    }
                    break;
                case nameof(CompletionByPeriod):
                    List<CompletionByPeriod> CompletionByPeriodList = new();
                    foreach (var row in ws.Rows(startRow, endRow))
                    {
                        if (!row.IsHidden)
                        {
                            row.Cell(Config.FinReportValues.CompletionByPeriod_PeriodName_Column).TryGetValue(out string periodName);
                            row.Cell(Config.FinReportValues.CompletionByPeriod_CurrentYear_Column).TryGetValue(out decimal currentYear);
                            row.Cell(Config.FinReportValues.CompletionByPeriod_PreviousYear_Column).TryGetValue(out decimal previousYear);
                            string? growthPercentage = GetPlanCompletionPercentageString(row.Cell(Config.FinReportValues.CompletionByPeriod_GrowthPercentage_Column));
                            row.Cell(Config.FinReportValues.CompletionByPeriod_BPCompletion_Column).TryGetValue(out decimal bPCompletion);
                            string? bPCompletionPercentage = GetPlanCompletionPercentageString(row.Cell(Config.FinReportValues.CompletionByPeriod_BPCompletionPercentage_Column));

                            CompletionByPeriod completionByPeriod = new()
                            {
                                PeriodName = periodName,
                                CurrentYear = currentYear,
                                PreviousYear = previousYear,
                                GrowthPercentage = growthPercentage,
                                BPCompletion = bPCompletion,
                                BPCompletionPercentage = bPCompletionPercentage
                            };
                            resultList.Add(completionByPeriod as T);
                        }
                    }
                    break;
                case nameof(Product):
                    List<Product> Products = new();
                    foreach (var row in ws.Rows(startRow, endRow))
                    {
                        if (!row.IsHidden)
                        {
                            row.Cell(Config.FinReportValues.Product_Name_Column).TryGetValue(out string name);
                            row.Cell(Config.FinReportValues.Product_PlanNumber_Column).TryGetValue(out int? planNumber);
                            row.Cell(Config.FinReportValues.Product_PlanPrice_Column).TryGetValue(out decimal planPrice);
                            row.Cell(Config.FinReportValues.Product_TotalPlanMoney_Column).TryGetValue(out decimal totalPlanMoney);
                            row.Cell(Config.FinReportValues.Product_FactNumber_Column).TryGetValue(out int? factNumber);
                            row.Cell(Config.FinReportValues.Product_TotalFactMoney_Column).TryGetValue(out decimal totalFactMoney);
                            string? planCompletionPercentage = GetPlanCompletionPercentageString(row.Cell(Config.FinReportValues.Product_PlanCompletionPercentage_Column));
                            row.Cell(Config.FinReportValues.Product_Type_Column).TryGetValue(out string? type);
                            row.Cell(Config.FinReportValues.Product_Plan_Column).TryGetValue(out int? plan);
                            row.Cell(Config.FinReportValues.Product_Fact_Column).TryGetValue(out int? fact);
                            row.Cell(Config.FinReportValues.Product_TotalPlanByToday_Column).TryGetValue(out int? totalPlanByToday);
                            row.Cell(Config.FinReportValues.Product_TotalFactByToday_Column).TryGetValue(out int? totalFactByToday);
                            row.Cell(Config.FinReportValues.Product_PlanDeviation_Column).TryGetValue(out int? planDeviation);

                            Product product = new()
                            {
                                Name = name,
                                PlanNumber = planNumber,
                                PlanPrice = Math.Round(planPrice, 2),
                                TotalPlanMoney = Math.Round(totalPlanMoney, 2),
                                FactNumber = factNumber,
                                TotalFactMoney = Math.Round(totalFactMoney, 2),
                                PlanCompletionPercentage = planCompletionPercentage,
                                Type = type,
                                Plan = plan,
                                Fact = fact,
                                TotalPlanByToday = totalPlanByToday,
                                TotalFactByToday = totalFactByToday,
                                PlanDeviation = planDeviation
                            };
                            resultList.Add(product as T);
                        }
                    }
                    break;
                default:
                    throw new NotImplementedException();
            }

            return resultList;
        }

        private string GetPlanCompletionPercentageString(IXLCell cell)
        {
            // пытаемся взять planCompletionPercentage точно в таком же виде, в каком он отображается в отчёте
            string planCompletionPercentage;
            if (!cell.TryGetValue(out decimal planCompletionPercentageDecimal))
            {
                cell.TryGetValue(out planCompletionPercentage);
            }
            else
            {
                planCompletionPercentage = Math.Round(planCompletionPercentageDecimal * 100, 1).ToString() + "%";
            }
            return planCompletionPercentage;
        }

        public void Dispose()
        {
            Wb.Dispose();
            this.Dispose();
        }
    }
}
