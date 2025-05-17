using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MPOVT_DataCenter_Visualisator.Models;

namespace MPOVT_DataCenter_Visualisator.Workers
{
    public class ConfigWorker
    {
        public Config GetConfig()
        {
            Config? config = Newtonsoft.Json.JsonConvert.DeserializeObject<Config>(File.ReadAllText("appsettings.json"));
            if (config is not null)
            {
                return config;
            }
            else
            {
                throw new Exception("Config is null");
            }
        }
    }
}
