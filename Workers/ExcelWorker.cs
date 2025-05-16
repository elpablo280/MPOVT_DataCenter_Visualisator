using System;
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
        private XLWorkbook? wb;
        private string _filepath;

        public ExcelWorker(string filepath)
        {
            _filepath = filepath;
        }

        public void OpenWorkbook()
        {
            wb = new(_filepath);
        }

        public void ImportData()
        {
            OpenWorkbook();
            if (wb is not null)
            {
                // Открываем Excel файл
                using (wb)
                {
                    var ws = wb.Worksheet(3);

                    ws.CollapseRows();
                    ws.CollapseColumns();
                    List<GeneralInfo> GeneralInfoList = new();
                    List<Product> Products = new();
                    List<CompletionByPeriod> CompletionByPeriodList = new();
                    try
                    {
                        GeneralInfoList = GetGeneralInfo(ws, 5, 25);
                        Products = GetProducts(ws, 31, 204);
                        CompletionByPeriodList = GetCompletionByPeriod(ws, 209, 234);
                    }
                    catch (Exception ex)
                    {

                    }

                }
            }
            else
            {

            }
        }

        private string GetPlanCompletionString(IXLCell cell)
        {
            // пытаемся взять planCompletionPercentage точно в таком же виде, в каком он отображается в отчёте
            bool IsPercentageDecimal = true;
            string? planCompletionPercentage;
            if (!cell.TryGetValue(out decimal planCompletionPercentageDecimal))
            {
                IsPercentageDecimal = false;
                cell.TryGetValue(out planCompletionPercentage);
            }
            else
            {
                planCompletionPercentage = Math.Round(planCompletionPercentageDecimal * 100, 1).ToString() + "%";
            }
            return planCompletionPercentage;
        }

        private List<GeneralInfo> GetGeneralInfo(IXLWorksheet ws, int startRow, int endRow)
        {
            List<GeneralInfo> GeneralInfoList = new();
            for (int i = startRow; i <= endRow; i++)
            {
                ws.Cell("B" + i).TryGetValue(out string name);
                ws.Cell("E" + i).TryGetValue(out decimal planValue);
                ws.Cell("F" + i).TryGetValue(out decimal factValue);
                string? planCompletionPercentage = GetPlanCompletionString(ws.Cell("G" + i));

                GeneralInfo generalInfo = new()
                {
                    Name = name,
                    PlanValue = Math.Round(planValue, 2),
                    FactValue = Math.Round(factValue, 2),
                    PlanCompletionPercentage = planCompletionPercentage,
                };
                GeneralInfoList.Add(generalInfo);
            }

            return GeneralInfoList;
        }

        private List<CompletionByPeriod> GetCompletionByPeriod(IXLWorksheet ws, int startRow, int endRow)
        {
            List<CompletionByPeriod> CompletionByPeriodList = new();
            for (int i = startRow; i <= endRow; i++)
            {
                ws.Cell("B" + i).TryGetValue(out string periodName);
                ws.Cell("C" + i).TryGetValue(out decimal currentYear);
                ws.Cell("D" + i).TryGetValue(out decimal previousYear);
                string? growthPercentage = GetPlanCompletionString(ws.Cell("E" + i));
                ws.Cell("F" + i).TryGetValue(out decimal bPCompletion);
                string? bPCompletionPercentage = GetPlanCompletionString(ws.Cell("G" + i));

                CompletionByPeriod completionByPeriod = new()
                {
                    PeriodName = periodName,
                    CurrentYear = currentYear,
                    PreviousYear = previousYear,
                    GrowthPercentage = growthPercentage,
                    BPCompletion = bPCompletion,
                    BPCompletionPercentage = bPCompletionPercentage
                };
                CompletionByPeriodList.Add(completionByPeriod);
            }

            return CompletionByPeriodList;
        }

        private List<Product> GetProducts(IXLWorksheet ws, int startRow, int endRow)
        {
            List<Product> Products = new();
            for (int i = startRow; i <= endRow; i++)
            {
                ws.Cell("B" + i).TryGetValue(out string name);
                ws.Cell("C" + i).TryGetValue(out int? planNumber);
                ws.Cell("D" + i).TryGetValue(out decimal planPrice);
                ws.Cell("E" + i).TryGetValue(out decimal totalPlanMoney);
                ws.Cell("F" + i).TryGetValue(out int? factNumber);
                ws.Cell("G" + i).TryGetValue(out decimal totalFactMoney);
                string? planCompletionPercentage = GetPlanCompletionString(ws.Cell("H" + i));
                ws.Cell("I" + i).TryGetValue(out string? type);
                ws.Cell("AJ" + i).TryGetValue(out int? plan);
                ws.Cell("AK" + i).TryGetValue(out int? fact);
                ws.Cell("BT" + i).TryGetValue(out int? totalPlanByToday);
                ws.Cell("BU" + i).TryGetValue(out int? totalFactByToday);
                ws.Cell("BV" + i).TryGetValue(out int? planDeviation);

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
                Products.Add(product);
            }

            return Products;
        }

        public void Dispose()
        {
            wb.Dispose();
            this.Dispose();
        }
    }
}
