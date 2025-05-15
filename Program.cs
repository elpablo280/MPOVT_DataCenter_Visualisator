using MPOVT_DataCenter_Visualisator.Jobs;
using MPOVT_DataCenter_Visualisator.Workers;
using Quartz.Spi;
using Serilog;

namespace MPOVT_DataCenter_Visualisator
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var cw = new ConfigWorker();
            var config = cw.GetConfig();

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.Console()
                .WriteTo.File(config.LogsFilepath, rollingInterval: RollingInterval.Day)
                .CreateLogger();
            try
            {
                Log.Information("Starting up the Worker Service");
                CreateHostBuilder(args).Build().Run();
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

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .UseSerilog()
                .ConfigureServices((hostContext, services) => { 
                    services.AddHostedService<Worker>();
                    services.AddSingleton<MainJob>();
                });
    }
}