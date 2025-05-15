using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;

namespace MPOVT_DataCenter_Visualisator.Workers
{
    public class GoogleSheetsWorker
    {
        public async Task ImportData()
        {
            // Путь к вашему JSON-файлу с учетными данными
            string credentialPath = "path/to/your/credentials.json";

            // Создание учетных данных
            GoogleCredential credential;
            using (var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }
            // Создание службы Google Sheets
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Google Sheets API Sample",
            });
            // ID таблицы и диапазон, куда вы хотите записать данные
            var spreadsheetId = "your_spreadsheet_id";
            var range = "Sheet1!A1:B2"; // Укажите нужный диапазон
                                        // Данные для записи
            var valueRange = new ValueRange
            {
                Values = new[] { new[] { "Hello", "World" }, new[] { "Foo", "Bar" } }
            };
            // Запись данных в таблицу
            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var response = await updateRequest.ExecuteAsync();
            Console.WriteLine("Данные успешно импортированы!");
        }
    }
}
