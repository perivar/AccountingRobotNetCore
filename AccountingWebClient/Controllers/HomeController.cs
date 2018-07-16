using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Security.Claims;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.SignalR;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using AccountingWebClient.Models;
using AccountingWebClient.Hubs;
using AccountingServices;
using System.Threading;

namespace AccountingWebClient.Controllers
{
    public class HomeController : Controller
    {
        // dependency injected in Startup.cs
        public IBackgroundTaskQueue Queue { get; }
        private readonly IApplicationLifetime _appLifetime;
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _appConfig;
        private readonly IHubContext<JobProgressHub> _hubContext;
        private readonly AccountingRobot _accountingRobot;


        public HomeController(IBackgroundTaskQueue queue,
            IApplicationLifetime appLifetime,
            ILogger<HomeController> logger,
            IConfiguration configuration,
            IHubContext<JobProgressHub> hubContext,
            AccountingRobot accountingRobot)
        {
            Queue = queue;
            _appLifetime = appLifetime;
            _logger = logger;
            _appConfig = configuration;
            _hubContext = hubContext;
            _accountingRobot = accountingRobot;
        }

        [AllowAnonymous]
        public IActionResult Login()
        {
            return View();
        }

        [AllowAnonymous, HttpPost]
        public IActionResult Login(LoginData loginData)
        {
            // get shopify configuration parameters
            string username = _appConfig["OberloUsername"];
            string password = _appConfig["OberloPassword"];

            if (ModelState.IsValid)
            {
                var isValid = (loginData.Username == username && loginData.Password == password);
                if (!isValid)
                {
                    ModelState.AddModelError("", "Login failed. Please check Username and/or password");
                    return View();
                }
                else
                {
                    var identity = new ClaimsIdentity(CookieAuthenticationDefaults.AuthenticationScheme, ClaimTypes.Name, ClaimTypes.Role);
                    identity.AddClaim(new Claim(ClaimTypes.NameIdentifier, loginData.Username));
                    identity.AddClaim(new Claim(ClaimTypes.Name, loginData.Username));
                    var principal = new ClaimsPrincipal(identity);
                    HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal, new AuthenticationProperties { IsPersistent = loginData.RememberMe });
                    return Redirect("~/Home");
                }
            }
            else
            {
                ModelState.AddModelError("", "username or password is blank");
                return View();
            }
        }

        [Authorize]
        public IActionResult Index()
        {
            return View();
        }


        [AllowAnonymous]
        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        [Authorize]
        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        [Authorize]
        public IActionResult StartProgress()
        {
            string jobId = Guid.NewGuid().ToString("N");

            Queue.QueueBackgroundWorkItem(cancellationToken => PerformBackgroundJob(jobId, cancellationToken));

            return RedirectToAction("Progress", new { jobId });
        }

        private async Task PerformBackgroundJob(string jobId, CancellationToken cancellationToken)
        {
            _logger.LogInformation(
                $"Queued Background Task {jobId} is running.");

            await _accountingRobot.DoProcessAsync(jobId, cancellationToken);

            _logger.LogInformation(
                $"Queued Background Task {jobId} is complete.");
        }

        public IActionResult Progress(string jobId)
        {
            ViewBag.JobId = jobId;

            return View();
        }
    }
}
