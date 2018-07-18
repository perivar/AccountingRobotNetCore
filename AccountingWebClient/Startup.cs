using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using AccountingWebClient.Hubs;
using AccountingServices;

namespace AccountingWebClient
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            // enable cookie authentication defaults
            services.AddAuthentication(options =>
            {
                options.DefaultSignInScheme = CookieAuthenticationDefaults.AuthenticationScheme;
                options.DefaultAuthenticateScheme = CookieAuthenticationDefaults.AuthenticationScheme;
                options.DefaultChallengeScheme = CookieAuthenticationDefaults.AuthenticationScheme;
            }).AddCookie(options => { options.LoginPath = "/Home/Login"; });

            // enable CORS access. Make sure this call is put just above app.AddMvc
            services.AddCors(options =>
            {
                options.AddPolicy("AllowAll",
                  builder =>
                  {
                      builder
                      .AllowAnyOrigin()
                      .AllowAnyMethod()
                      .AllowAnyHeader()
                      .AllowCredentials();
                  });

                options.AddPolicy("AllowLocalhost",
                    builder =>
                    {
                        builder
                        .AllowAnyMethod()
                        .AllowAnyHeader()
                        .WithOrigins("http://localhost:5001");
                    });
            });

            // ensure the logon page is anonymous
            services.AddMvc()
            .AddRazorPagesOptions(options =>
            {
                options.Conventions.AuthorizeFolder("/");
                options.Conventions.AllowAnonymousToPage("/Home/Login");
            })
            .SetCompatibilityVersion(CompatibilityVersion.Version_2_1);

            // Adds a default in-memory implementation of IDistributedCache.
            services.AddDistributedMemoryCache();

            // Enable session storage
            services.AddSession();

            // Enable SignalR
            services.AddSignalR();

            // Enable the background services
            services.AddSingleton<AccountingRobot>();
            services.AddHostedService<QueuedHostedService>();
            services.AddSingleton<IBackgroundTaskQueue, BackgroundTaskQueue>();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, Microsoft.AspNetCore.Hosting.IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
                app.UseHsts();
            }

            // add CORS for signalr to work. Make sure this call is put just above app.UseSignalR
            app.UseCors("AllowLocalhost");

            // add signalr hub url
            app.UseSignalR(routes =>
            {
                routes.MapHub<JobProgressHub>("/jobprogress");
            });

            // app.UseHttpsRedirection();

            // serve static files in the wwwroot folder
            app.UseStaticFiles();

            // enable authentication
            app.UseAuthentication();

            // IMPORTANT: This session call MUST go before UseMvc()
            app.UseSession();

            // the default route is sufficient
            // {controller=Home}/{action=Index}/{id?}
            app.UseMvcWithDefaultRoute();

        }
    }
}
