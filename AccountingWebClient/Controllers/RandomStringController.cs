using System;
using Microsoft.AspNetCore.Mvc;

namespace AccountingWebClient.Controllers
{
    [Route("api/[controller]")]
    public class RandomStringController : Controller
    {
        private readonly RandomStringProvider _randomStringProvider;

        public RandomStringController(RandomStringProvider randomStringProvider)
        {
            _randomStringProvider = randomStringProvider;
        }

        [HttpGet]
        public string Get()
        {
            return _randomStringProvider.RandomString;
        }
    }

}
