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
            var environment = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");

            if (String.IsNullOrWhiteSpace(environment))
            {
                Console.WriteLine("Environment variable ASPNETCORE_ENVIRONMENT not set!");
                Console.WriteLine("Using appsettings.json configuration file."); 
            }
            else
            {
                Console.WriteLine("Environment: {0}", environment);
            }

            var configurationBuilder = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            if (environment == "Development")
            {
                configurationBuilder
                    .AddJsonFile($"appsettings.{environment}.json", optional: false);
            }
            configurationBuilder
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