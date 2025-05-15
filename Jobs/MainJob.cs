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

        public Task Execute(IJobExecutionContext context)
        {



            _logger.LogInformation("Задача выполнена в: {time}", DateTime.Now);
            return Task.CompletedTask;
        }
    }
}
