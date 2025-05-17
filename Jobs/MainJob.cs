using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MPOVT_DataCenter_Visualisator.Models;
using MPOVT_DataCenter_Visualisator.Workers;
using Quartz;
using static Quartz.Logging.OperationName;

namespace MPOVT_DataCenter_Visualisator.Jobs
{
    public class MainJob : IJob
    {
        private ILogger<MainJob> _logger;

        public MainJob(ILogger<MainJob> logger)
        {
            _logger = logger;
        }

        public Task Execute(IJobExecutionContext context)
        {
            _logger.LogInformation("Начало работы");
            // получаем данные из конфиг-файла
            Config? config = null;
            try
            {
                ConfigWorker cw = new();
                config = cw.GetConfig();
            }
            catch (Exception ex)
            {
                _logger.LogError("Возникла ошибка при получении данных из конфигурационного файла: " + ex.Message);
            }
            string FinReportFilepath = Path.Combine(config.DataCenterFolderpath, config.ProductionMonitoringFilepath);
            string FinReportFilepath_Copy = Path.Combine(config.DataCenterFolderpath, "Copy_" + config.ProductionMonitoringFilepath);

            // работа с экономическим отчётом
            ExcelWorker ew = new(FinReportFilepath, config);

            List<GeneralInfo> GeneralInfoList = new();
            List<Product> Products = new();
            List<CompletionByPeriod> CompletionByPeriodList = new();

            try
            {
                (GeneralInfoList, Products, CompletionByPeriodList) = ew.ImportData();
                _logger.LogInformation("Данные из файла мониторинга производства получены успешно");
            }
            catch (Exception ex)
            {
                _logger.LogError("Возникла ошибка при обработке файла мониторинга производства: " + ex.Message);
            }
            finally
            {
                //File.Delete(FinReportFilepath_Copy);
                ew.Dispose();
            }

            // экспорт данных в гугл таблицы todo


            _logger.LogInformation("Задача выполнена в: {time}", DateTime.Now);
            return Task.CompletedTask;
        }
    }
}
