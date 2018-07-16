using Microsoft.AspNetCore.SignalR;
using System.Threading.Tasks;

namespace AccountingWebClient.Hubs
{
    public class JobProgressHub : Hub
    {
        public async Task AssociateJob(string jobId)
        {
            await Groups.AddToGroupAsync(Context.ConnectionId, jobId);
            await Clients.All.SendAsync("ReceiveMessage", "HUB", "job associated with " + jobId);
        }

        public async Task SendMessage(string user, string message)
        {
            await Clients.All.SendAsync("ReceiveMessage", user, message);
        }

        public async Task SendJobMessage(string jobId, string user, string message)
        {
            await Clients.Group(jobId).SendAsync("ReceiveMessage", user, message);
        }
    }
}