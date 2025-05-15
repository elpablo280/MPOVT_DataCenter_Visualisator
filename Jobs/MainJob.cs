using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quartz;
using static Quartz.Logging.OperationName;

namespace MPOVT_DataCenter_Visualisator.Jobs
{
    public class MainJob : IJob
    {
        private readonly ILogger<MainJob> _logger;

        public MainJob(ILogger<MainJob> logger)
        {
            _logger = logger;
        }

        //public MainJob()
        //{

        //}

        public Task Execute(IJobExecutionContext context)
        {
            //ILogger<MainJob> _logger;
            // Логика вашей задачи
            _logger.LogInformation("Задача выполнена в: {time}", DateTime.Now);
            return Task.CompletedTask;
        }
    }
}
