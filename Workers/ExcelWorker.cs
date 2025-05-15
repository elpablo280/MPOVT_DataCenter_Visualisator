using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace MPOVT_DataCenter_Visualisator.Workers
{
    public class ExcelWorker(string filepath)
    {
        XLWorkbook? workbook = new(filepath);

        public void ImportData()
        {
            // Открываем Excel файл
            using (workbook)
            {
                // Получаем второй лист
                var worksheet = workbook.Worksheet(2);
                // Получаем количество строк и столбцов
                int rowCount = worksheet.LastRowUsed().RowNumber();
                int columnCount = worksheet.LastColumnUsed().ColumnNumber();
                // Читаем данные из ячеек
                //for (int row = 1; row <= rowCount; row++)
                //{
                //    for (int col = 1; col <= columnCount; col++)
                //    {
                //        var cellValue = worksheet.Cell(row, col).Value;
                //        Console.Write($"{cellValue} ");
                //    }
                //    Console.WriteLine();
                //}
            }
        }
    }
}
