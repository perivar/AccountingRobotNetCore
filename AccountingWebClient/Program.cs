using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace AccountingWebClient
{
   public class Program
    {
        public static void Main(string[] args)
        {
            BuildWebHost(args).Run();
        }

        public static IWebHost BuildWebHost(string[] args) =>

            /*
            The tasks that are performed by CreateDefaultBuilder is
            1. Configures Kestrel as the web server.
            2. Sets the content root to Directory.GetCurrentDirectory.
            3. Loads optional configuration from
                a) Appsettings.json
                b) Appsettings.{Environment}.json.
                c) User secrets when the app runs in the Development environment.
                d) Environment variables
                e) Command-line arguments.
            4. Enable logging
            5. Integrates the Kestrel run with IIS             
             */             
            WebHost.CreateDefaultBuilder(args)
                .ConfigureAppConfiguration(builder =>
                {
                    // Add UserSecrets for both development and Production environment
                    builder.AddUserSecrets<Startup>();
                })
                .UseStartup<Startup>()
                .Build();
    }
}
