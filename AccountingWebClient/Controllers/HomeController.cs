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
        private readonly RandomStringProvider _randomStringProvider;

        public HomeController(IBackgroundTaskQueue queue,
            IApplicationLifetime appLifetime,
            ILogger<HomeController> logger,
            IConfiguration configuration,
            IHubContext<JobProgressHub> hubContext,
            RandomStringProvider randomStringProvider)
        {
            Queue = queue;
            _appLifetime = appLifetime;
            _logger = logger;
            _appConfig = configuration;
            _hubContext = hubContext;
            _randomStringProvider = randomStringProvider;
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

        [Authorize]
        public IActionResult Process()
        {
            AccountingRobot.Program.Process();

            ViewData["Message"] = "Accounting Spreadsheet Updated";

            return View("Index");
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
        public IActionResult StartProgress()
        {
            string jobId = Guid.NewGuid().ToString("N");

            Queue.QueueBackgroundWorkItem(async cancellationToken =>
                {
                    _logger.LogInformation(
                        $"Queued Background Task {jobId} is running.");


                    for (int delayLoop = 0; delayLoop < 10; delayLoop++)
                    {
                        _logger.LogInformation(
                            $"Queued Background Task {jobId} is running. {delayLoop}/10");

                        await _randomStringProvider.UpdateString(cancellationToken);
                        await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
                    }

                    _logger.LogInformation(
                        $"Queued Background Task {jobId} is complete. 10/10");
                });

            return RedirectToAction("Progress", new { jobId });
        }

        public IActionResult Progress(string jobId)
        {
            ViewBag.JobId = jobId;

            return View();
        }
    }
}
