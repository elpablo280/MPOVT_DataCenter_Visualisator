using Quartz.Impl;
using Quartz;
using MPOVT_DataCenter_Visualisator.Jobs;
using MPOVT_DataCenter_Visualisator.Models;
using MPOVT_DataCenter_Visualisator.Workers;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace MPOVT_DataCenter_Visualisator
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            ConfigWorker cw = new();
            var config = cw.GetConfig();

            // Создание планировщика
            var schedulerFactory = new StdSchedulerFactory();
            var scheduler = await schedulerFactory.GetScheduler(stoppingToken);
            await scheduler.Start(stoppingToken);
            // Создание задания
            var job = JobBuilder.Create<MainJob>()
                .WithIdentity("mainJob", "group1")
                .Build();
            // Создание триггера с использованием крон-выражения
            var trigger = TriggerBuilder.Create()
                .WithIdentity("myTrigger", "group1")
                .WithCronSchedule(config.CronExpression) // Использование крон-выражения
                .Build();
            // Запуск задания
            await scheduler.ScheduleJob(job, trigger, stoppingToken);



            //while (!stoppingToken.IsCancellationRequested)
            //{
            //    if (_logger.IsEnabled(LogLevel.Information))
            //    {
            //        _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
            //    }
            //    await Task.Delay(1000, stoppingToken);
            //}
        }
    }
}
