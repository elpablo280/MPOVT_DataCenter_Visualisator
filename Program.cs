using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using MPOVT_DataCenter_Visualisator;
using MPOVT_DataCenter_Visualisator.Workers;
using Serilog;
using Quartz.Impl;
using Quartz;
using MPOVT_DataCenter_Visualisator.Models;
using MPOVT_DataCenter_Visualisator.Jobs;

public class Program
{
    public static void Main(string[] args)
    {
        ConfigWorker cw = new();
        var config = cw.GetConfig();

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.Console()
            .WriteTo.File(config.LogsFilepath, rollingInterval: RollingInterval.Day)
            .CreateLogger();

        try
        {
            Log.Information("Starting up the host");
            CreateHostBuilder(args, config).Build().Run();
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Application start-up failed");
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }

    public static IHostBuilder CreateHostBuilder(string[] args, Config config) =>
        Host.CreateDefaultBuilder(args)
            .UseSerilog() // Используем Serilog
            .ConfigureServices((hostContext, services) =>
            {
                services.AddQuartz(q =>
                {
                    q.UseMicrosoftDependencyInjectionJobFactory();

                    // Создаем задачу
                    var jobKey = new JobKey("MainJob");
                    q.AddJob<MainJob>(opts => opts.WithIdentity(jobKey));

                    // Создаем триггер
                    q.AddTrigger(opts => opts
                        .ForJob(jobKey)
                        .WithIdentity("MainJob-trigger")
                        .WithCronSchedule(config.CronExpression));
                });
                services.AddQuartzHostedService(q => q.WaitForJobsToComplete = true);
            });
}
