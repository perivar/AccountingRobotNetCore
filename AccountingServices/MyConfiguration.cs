using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace AccountingServices
{
    public class MyConfiguration : IMyConfiguration
    {
        public static IConfiguration Configuration { get; set; }

        public MyConfiguration()
        {

            var environmentName = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");

            var configurationBuilder = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .AddJsonFile($"appsettings.{environmentName}.json", optional: true)
                .AddUserSecrets<MyConfiguration>()
                .AddEnvironmentVariables();

            Configuration = configurationBuilder.Build();
        }

        public string GetValue(string key)
        {
            return Configuration[key];
        }
    }
}