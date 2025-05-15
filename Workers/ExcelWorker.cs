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
                    List<Product> products = new();
                    try
                    {
                        products = GetProducts(ws, 31, 204);

                    }
                    catch (Exception ex)
                    {
                    }

                    IXLCell startCell = ws.Cell("B5");
                    IXLCell endCell = ws.Cell("G22");
                    List<GeneralInfo> generals = [];
                    int startRow = 5;
                    int endRow = 25;
                    int startRow1 = 26;
                    int endRow1 = 204;
                    for (int i = startRow; i <= endRow; i++)
                    {
                        
                        var value1 = ws.Cell("B" + i).GetString();
                        var value2 = ws.Cell("E" + i).GetString();
                        var value3 = ws.Cell("F" + i).GetString();
                        var value4 = ws.Cell("G" + i).GetString();
                        GeneralInfo generalInfo = new GeneralInfo(
                            ws.Cell("B" + i).GetString(),
                            Math.Round(Convert.ToDouble(ws.Cell("E" + i).GetString()), 2),
                            Math.Round(Convert.ToDouble(ws.Cell("F" + i).GetString()), 2),
                            ws.Cell("G" + i).GetString()
                            );
                        generals.Add(generalInfo);
                    }

                    // Получаем количество строк и столбцов
                    int rowCount = ws.LastRowUsed().RowNumber();
                    int columnCount = ws.LastColumnUsed().ColumnNumber();
                }
            }
            else
            {

            }
        }

        public List<Product> GetProducts(IXLWorksheet ws, int startRow, int endRow)
        {
            List<Product> Products = new List<Product>();
            string name;
            int? planNumber;
            decimal planPrice;
            decimal totalPlanMoney;
            int? factNumber;
            decimal totalFactMoney;
            string? planCompletionPercentage;
            string? type;
            int? plan;
            int? fact;
            int? totalPlanByToday;
            int? totalFactByToday;
            int? planDeviation;
            for (int i = startRow; i <= endRow; i++)
            {
                ws.Cell("B" + i).TryGetValue(out name);
                ws.Cell("C" + i).TryGetValue(out planNumber);
                ws.Cell("D" + i).TryGetValue(out planPrice);
                ws.Cell("E" + i).TryGetValue(out totalPlanMoney);
                ws.Cell("F" + i).TryGetValue(out factNumber);
                ws.Cell("G" + i).TryGetValue(out totalFactMoney);
                ws.Cell("H" + i).TryGetValue(out planCompletionPercentage);
                ws.Cell("I" + i).TryGetValue(out type);
                ws.Cell("AJ" + i).TryGetValue(out plan);
                ws.Cell("AK" + i).TryGetValue(out fact);
                ws.Cell("BT" + i).TryGetValue(out totalPlanByToday);
                ws.Cell("BU" + i).TryGetValue(out totalFactByToday);
                ws.Cell("BV" + i).TryGetValue(out planDeviation);

                Product product = new(
                    name,
                    planNumber,
                    Math.Round(planPrice, 2),
                    Math.Round(totalPlanMoney, 2),
                    factNumber,
                    Math.Round(totalFactMoney, 2),
                    planCompletionPercentage,
                    type,
                    plan,
                    fact,
                    totalPlanByToday,
                    totalFactByToday,
                    planDeviation);
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
