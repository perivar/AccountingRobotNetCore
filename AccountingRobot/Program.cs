using System;
using System.Threading;
using AccountingServices;

namespace AccountingRobotCLI
{
    public class Program
    {
        static void Main(string[] args)
        {
            var accountingRobot = new AccountingRobot();
            var s = new CancellationTokenSource();
            string jobId = Guid.NewGuid().ToString("N");

            accountingRobot.DoProcessAsync(jobId, s.Token).GetAwaiter();
        }
    }
}
