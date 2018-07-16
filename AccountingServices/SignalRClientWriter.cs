using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.SignalR.Client;
using Microsoft.Extensions.Logging;

namespace AccountingServices
{
    public class SignalRClientWriter : TextWriter
    {
        HubConnection HubConnection { get; set; }
        public string JobId { get; set; }
        bool HubConnectionStarted { get; set; }
        bool FlushAfterEveryWrite { get; set; }

        public SignalRClientWriter(string url) : this(url, null)
        {
        }

        public SignalRClientWriter(string url, string jobId)
        {
            JobId = jobId;
            HubConnectionStarted = false;
            FlushAfterEveryWrite = false;

            // https://docs.microsoft.com/en-us/aspnet/core/signalr/configuration?view=aspnetcore-2.1
            HubConnection = new HubConnectionBuilder()
            .WithUrl(url)
            .ConfigureLogging(logging =>
            {
                logging.SetMinimumLevel(LogLevel.Information);
                logging.AddConsole();
            })
            .Build();

            // open connection
            CheckOrOpenConnection().GetAwaiter();
        }

        public override async Task WriteAsync(string value)
        {
            await OutputMessage(value);

            if (FlushAfterEveryWrite)
                await FlushAsync();
        }

        public override async Task WriteLineAsync(string value)
        {
            await OutputMessage(value);

            if (FlushAfterEveryWrite)
                await FlushAsync();
        }

        public override async Task WriteLineAsync()
        {
            await OutputMessage(null);

            if (FlushAfterEveryWrite)
                await FlushAsync();
        }

        public override async Task FlushAsync()
        {
            // do nothing
        }

        public override void Write(string value)
        {
            OutputMessage(value).GetAwaiter();

            if (FlushAfterEveryWrite)
                Flush();
        }

        public override void WriteLine(string value)
        {
            OutputMessage(value).GetAwaiter();

            if (FlushAfterEveryWrite)
                Flush();
        }

        public override void WriteLine()
        {
            OutputMessage(null).GetAwaiter();

            if (FlushAfterEveryWrite)
                Flush();
        }

        public override void Flush()
        {
        }

        public override Encoding Encoding => throw new System.NotImplementedException();

        private async Task CheckOrOpenConnection()
        {
            if (!HubConnectionStarted)
            {
                try
                {
                    await HubConnection.StartAsync();
                    HubConnectionStarted = true;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("Failed starting SignalR client: {0}", ex.Message);
                }
            }
        }

        private async Task OutputMessage(string message)
        {
            await CheckOrOpenConnection();

            if (HubConnectionStarted)
            {
                try
                {
                    await HubConnection.InvokeAsync("SendJobMessage", JobId, "Robot", message);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("Failed sending message to SignalR Hub: {0}", ex.Message);
                }
            }

            Console.WriteLine(message);
        }
    }
}