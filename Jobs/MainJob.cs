using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            ConfigWorker cw = new();
            var config = cw.GetConfig();
            string FinReportFilepath = Path.Combine(config.DataCenterFolderpath, config.ProductionMonitoringFilepath);
            string FinReportFilepath_Copy = Path.Combine(config.DataCenterFolderpath, "Copy_" + config.ProductionMonitoringFilepath);

            // todo
            //if (File.Exists(FinReportFilepath))
            //{
            //    File.Copy(FinReportFilepath, FinReportFilepath_Copy);
            //}

            // работа с экономическим отчётом
            ExcelWorker ew = new(FinReportFilepath);
            try
            {
                ew.ImportData();

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

            _logger.LogInformation("Задача выполнена в: {time}", DateTime.Now);
            return Task.CompletedTask;
        }
    }
}
