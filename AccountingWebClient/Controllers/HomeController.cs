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
using Microsoft.Extensions.Configuration;
using AccountingWebClient.Models;

namespace AccountingWebClient.Controllers
{
    public class HomeController : Controller
    {
        private readonly IConfiguration appConfig;

        public HomeController(IConfiguration configuration)
        {
            appConfig = configuration;
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
            string username = appConfig["OberloUsername"];
            string password = appConfig["OberloPassword"];

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
    }
}
