using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using AccountingWebClient.Hubs;
using Microsoft.AspNetCore.SignalR;

namespace AccountingWebClient
{
    public class RandomStringProvider
    {
        // Inject an instance of IHubContext<JobProgressHub> into the controller and then call SendAsync for the group.
        private readonly IHubContext<JobProgressHub> _hubContext;

        private const string RandomStringUri =
                "https://www.random.org/strings/?num=1&len=8&digits=on&upperalpha=on&loweralpha=on&unique=on&format=plain&rnd=new";

        private readonly HttpClient _httpClient;
        private int counter = 0;

        public RandomStringProvider(IHubContext<JobProgressHub> hubContext)
        {
            _httpClient = new HttpClient();
            _hubContext = hubContext;
        }

        public async Task UpdateString(CancellationToken cancellationToken)
        {
            try
            {
                var response = await _httpClient.GetAsync(RandomStringUri, cancellationToken);
                await _hubContext.Clients.All.SendAsync("progress", counter++);

                if (response.IsSuccessStatusCode)
                {
                    RandomString = await response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public string RandomString { get; private set; } = string.Empty;
    }
}
