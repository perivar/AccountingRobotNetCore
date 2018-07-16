using System;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.Reflection;
using Microsoft.Extensions.Configuration;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using AccountingServices.Excel;
using AccountingServices.Helpers;

namespace AccountingServices.SBankenService
{
    public static class SBanken
    {
        public static SkandiabankenBankStatement GetBankStatementFromTransactions(List<SBankenTransaction> sBankenTransactions)
        {
            var date = new Date();
            var firstDayOfTheYear = date.FirstDayOfTheYear;

            var bankStatment = new SkandiabankenBankStatement
            {
                Transactions = sBankenTransactions,
                IncomingBalanceDate = firstDayOfTheYear,
                IncomingBalanceLabel = string.Format("INNGÅENDE SALDO {0:dd.MM.yyyy}", firstDayOfTheYear),
                IncomingBalance = 0,
                OutgoingBalanceDate = DateTime.MinValue,
                OutgoingBalanceLabel = null,
                OutgoingBalance = 0
            };

            return bankStatment;
        }

        public static SkandiabankenBankStatement GetLatestBankStatement(IMyConfiguration configuration, bool forceUpdate = false)
        {
            string cacheDir = configuration.GetValue("CacheDir");
            string cacheFileNamePrefix = configuration.GetValue("SBankenAccountNumber");

            string dateFromToRegexPattern = @"(\d{4}_\d{2}_\d{2})\-(\d{4}_\d{2}_\d{2})\.xlsx$";
            var lastCacheFileInfo = Utils.FindLastCacheFile(cacheDir, cacheFileNamePrefix, dateFromToRegexPattern, "yyyy_MM_dd", "_");

            var date = new Date();
            var currentDate = date.CurrentDate;
            var firstDayOfTheYear = date.FirstDayOfTheYear;
            var yesterday = date.Yesterday;

            // check if we have a cache file
            DateTime from = default(DateTime);
            DateTime to = default(DateTime);

            // if the cache file object has values
            if (!lastCacheFileInfo.Equals(default(KeyValuePair<DateTime, string>)))
            {
                from = lastCacheFileInfo.To;
                to = yesterday;

                // if the from date is today, then we already have an updated file so use cache
                if (from.Date.Equals(to.Date))
                {
                    // use latest cache file
                    return ReadBankStatement(lastCacheFileInfo.FilePath);
                }
            }

            // Download all from beginning of year until now
            from = firstDayOfTheYear;
            to = yesterday;

            // check special case if yesterday is last year
            if (from.Year != to.Year)
            {
                from = from.AddYears(-1);
            }

            // get updated bank statement
            string bankStatementFilePath = DownloadBankStatement(configuration, from, to);
            return ReadBankStatement(bankStatementFilePath);
        }

        public static SkandiabankenBankStatement ReadBankStatement(string skandiabankenTransactionsFilePath)
        {
            var skandiabankenTransactions = new List<SBankenTransaction>();

            var wb = new XLWorkbook(skandiabankenTransactionsFilePath);
            var ws = wb.Worksheet("Kontoutskrift");

            var startColumn = ws.Column(1);
            var firstCellFirstColumn = startColumn.FirstCellUsed();
            var lastCellFirstColumn = startColumn.LastCellUsed();
            var lastCellLastColumn = lastCellFirstColumn.WorksheetRow().AsRange().LastColumnUsed().LastCellUsed();

            // check edge case where first cell and last cell is the same (i.e. the spreadsheet contain no data)
            if (firstCellFirstColumn == lastCellFirstColumn)
            {
                // the spreadsheet contain no data
                Console.Out.WriteLine("ERROR! Bank statement contains no data.");
                return null;
            }

            // Get a range with the transaction data
            var transactionRange = ws.Range(firstCellFirstColumn, lastCellLastColumn).RangeUsed();

            // Treat the range as a table
            var transactionTable = transactionRange.AsTable();

            // Get the transactions
            foreach (var row in transactionTable.DataRange.Rows())
            {
                // BOKFØRINGSDATO	
                // RENTEDATO	
                // ARKIVREFERANSE	
                // TYPE	
                // TEKST	
                // UT FRA KONTO	
                // INN PÅ KONTO
                var skandiabankenTransaction = new SBankenTransaction();
                skandiabankenTransaction.TransactionDate = row.Field(0).GetDateTime();
                skandiabankenTransaction.InterestDate = row.Field(1).GetDateTime();
                skandiabankenTransaction.ArchiveReference = row.Field(2).GetString();
                skandiabankenTransaction.Type = row.Field(3).GetString();
                skandiabankenTransaction.Text = row.Field(4).GetString();
                skandiabankenTransaction.OutAccount = row.Field(5).GetValue<decimal>();
                skandiabankenTransaction.InAccount = row.Field(6).GetValue<decimal>();

                // set account change
                decimal accountChange = skandiabankenTransaction.InAccount - skandiabankenTransaction.OutAccount; ;
                skandiabankenTransaction.AccountChange = accountChange;

                if (accountChange > 0)
                {
                    skandiabankenTransaction.AccountingType = SBankenTransaction.AccountingTypeEnum.IncomeUnknown;
                }
                else
                {
                    skandiabankenTransaction.AccountingType = SBankenTransaction.AccountingTypeEnum.CostUnknown;
                }

                skandiabankenTransactions.Add(skandiabankenTransaction);
            }

            // find the incoming and outgoing balance
            var incomingBalanceCell = ws.Cell(lastCellLastColumn.Address.RowNumber + 2, lastCellLastColumn.Address.ColumnNumber);
            var outgoingBalanceCell = ws.Cell(1, lastCellLastColumn.Address.ColumnNumber);
            decimal incomingBalance = incomingBalanceCell.GetValue<decimal>();
            decimal outgoingBalance = outgoingBalanceCell.GetValue<decimal>();
            var incomingBalanceLabelCell = ws.Cell(lastCellLastColumn.Address.RowNumber + 2, lastCellLastColumn.Address.ColumnNumber - 2);
            var outgoingBalanceLabelCell = ws.Cell(1, lastCellLastColumn.Address.ColumnNumber - 2);
            var incomingBalanceLabel = incomingBalanceLabelCell.GetString();
            var outgoingBalanceLabel = outgoingBalanceLabelCell.GetString();
            var incomingBalanceDate = ExcelUtils.GetDateFromBankStatementString(incomingBalanceLabel);
            var outgoingBalanceDate = ExcelUtils.GetDateFromBankStatementString(outgoingBalanceLabel);

            var bankStatment = new SkandiabankenBankStatement
            {
                Transactions = skandiabankenTransactions,
                IncomingBalanceDate = incomingBalanceDate,
                IncomingBalanceLabel = incomingBalanceLabel,
                IncomingBalance = incomingBalance,
                OutgoingBalanceDate = outgoingBalanceDate,
                OutgoingBalanceLabel = outgoingBalanceLabel,
                OutgoingBalance = outgoingBalance
            };

            return bankStatment;
        }

        public static string DownloadBankStatement(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            string cacheDir = configuration.GetValue("CacheDir");
            string userDataDir = configuration.GetValue("UserDataDir");
            string sbankenMobilePhone = configuration.GetValue("SBankenMobilePhone");
            string sbankenBirthDate = configuration.GetValue("SBankenBirthDate");
            string cacheFileNamePrefix = configuration.GetValue("SBankenAccountNumber");
            string sbankenAccountId = configuration.GetValue("SBankenAccountId");
            string downloadFolderPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\Downloads\";
            string chromeDriverExePath = configuration.GetValue("ChromeDriverExePath");

            var driver = Utils.GetChromeWebDriver(userDataDir, chromeDriverExePath);

            driver.Navigate().GoToUrl("https://secure.sbanken.no/Authentication/BankIdMobile");

            var waitLoginPage = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            waitLoginPage.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            // login if login form is present
            if (SeleniumUtils.IsElementPresent(driver, By.Id("MobilePhone")))
            {
                // https://secure.sbanken.no/Authentication/BankIDMobile
                // input id MobilePhone - Mobilnummer (8 siffer)
                // input it BirthDate - Fødselsdato (ddmmåå)
                // submit value = Neste

                // login if login form is present
                if (SeleniumUtils.IsElementPresent(driver, By.XPath("//input[@id='MobilePhone']"))
                    && SeleniumUtils.IsElementPresent(driver, By.XPath("//input[@id='BirthDate']")))
                {
                    IWebElement mobilePhone = driver.FindElement(By.XPath("//input[@id='MobilePhone']"));
                    IWebElement birthDate = driver.FindElement(By.XPath("//input[@id='BirthDate']"));

                    mobilePhone.Clear();
                    mobilePhone.SendKeys(sbankenMobilePhone);

                    birthDate.Clear();
                    birthDate.SendKeys(sbankenBirthDate);

                    // use birth date field to submit form
                    birthDate.Submit();

                    var waitLoginIFrame = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                    waitLoginIFrame.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
                }
            }

            try
            {
                var waitMainPage = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                waitMainPage.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.UrlToBe("https://secure.sbanken.no/Home/Overview/Full#/"));
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Timeout - Logged in to Skandiabanken too late. Stopping.");
                return null;
            }

            // download account statement
            string accountStatementDownload = string.Format("https://secure.sbanken.no/Home/AccountStatement/ViewExcel?AccountId={0}&CustomFromDate={1:dd.MM.yyyy}&CustomToDate={2:dd.MM.yyyy}&FromDate=CustomPeriod&Incoming=", sbankenAccountId, from, to);
            driver.Navigate().GoToUrl(accountStatementDownload);

            var waitExcel = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            waitExcel.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            string accountStatementFileName = string.Format("{0}_{1:yyyy_MM_dd}-{2:yyyy_MM_dd}.xlsx", cacheFileNamePrefix, from, to);
            string accountStatementDownloadFilePath = Path.Combine(downloadFolderPath, accountStatementFileName);

            // wait until file has downloaded
            for (var i = 0; i < 30; i++)
            {
                if (File.Exists(accountStatementDownloadFilePath)) { break; }
                Thread.Sleep(1000);
            }
            var length = new FileInfo(accountStatementDownloadFilePath).Length;
            for (var i = 0; i < 30; i++)
            {
                Thread.Sleep(1000);
                var newLength = new FileInfo(accountStatementDownloadFilePath).Length;
                if (newLength == length && length != 0) { break; }
                length = newLength;
            }
            driver.Close();

            // determine path
            Console.Out.WriteLine("Successfully downloaded skandiabanken account statement excel file {0}", accountStatementDownloadFilePath);

            // moving file to right place
            string accountStatementDestinationPath = Path.Combine(cacheDir, accountStatementFileName);

            // To copy a folder's contents to a new location:
            // Create a new target folder, if necessary.
            if (!Directory.Exists(cacheDir))
            {
                Directory.CreateDirectory(cacheDir);
            }

            // Move file to another location
            File.Move(accountStatementDownloadFilePath, accountStatementDestinationPath);

            return accountStatementDestinationPath;
        }
    }

    public class SkandiabankenBankStatement
    {
        public List<SBankenTransaction> Transactions { get; set; }
        public DateTime IncomingBalanceDate { get; set; }
        public string IncomingBalanceLabel { get; set; }
        public decimal IncomingBalance { get; set; }
        public DateTime OutgoingBalanceDate { get; set; }
        public string OutgoingBalanceLabel { get; set; }
        public decimal OutgoingBalance { get; set; }
    }
}
