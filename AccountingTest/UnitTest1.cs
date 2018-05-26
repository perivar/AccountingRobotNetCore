using System;
using Xunit;
using Serilog;
using AccountingServices;

namespace AccountingTest
{
    public class UnitTest1
    {
        [Fact]
        public void TestCreateMD5()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.File("consoleapp.log")
                .CreateLogger();

            string input = "18.05.2018 00.00.0018.05.2018 00.00.00VISA VARE*2591 13.05 USD 41.38 SHOPIFY * 50900801 Kurs: 8.3038-343,61";
            string output = Utils.CreateMD5(input);
            Log.Information("{0}={1}", input, output);
            Assert.Equal("5e725a0458a0b4d7fed42bbf2f1495d1", output, StringComparer.CurrentCultureIgnoreCase);

        }
    }
}
