using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace AccountingServices
{
    public class SBankenTransaction
    {
        public enum AccountingTypeEnum
        {
            CostOfGoods,
            CostOfAdvertising,
            CostOfWebShop,
            CostOfDomain,
            CostOfServer,
            CostOfBank,
            CostOfInvoice,
            CostOfTryouts,
            CostUnknown,
            TransferStripe,
            TransferPaypal,
            TransferVipps,
            TransferUnknown,
            IncomeUnknown,
            IncomeReturn,
            IncomeInterest
        };

        // match e.g. *2591 29.04 NOK 159.66 PAYPAL *WANGFENG Kurs: 1.0000
        private static Regex purchasePattern = new Regex(@"(\*\d{4})\s(\d+\.\d+)\s(\w+)\s(\d+\.\d+)\s([\w\.\*\s]+)\s(Kurs\:)\s(\d+\.\d+)", RegexOptions.Compiled);

        // match e.g. Fra: The Currency Cloud Ltd Betalt: 24.08.17
        private static Regex transferPattern = new Regex(@"Fra\:\s([\w\s]+)\sBetalt\:\s(\d+\.\d+.\d+)", RegexOptions.Compiled);

        // match e.g. STRIPE PAYMENTS UK LTD             NOK             874,45
        private static Regex transferPattern2 = new Regex(@"([\w\s]+)\s+(\w{3})\s+(\d+[\.,]\d+)", RegexOptions.Compiled);

        public DateTime TransactionDate { get; set; }
        public DateTime InterestDate { get; set; }
        public string ArchiveReference { get; set; }
        public string Type { get; set; } // Overføring (intern), Overførsel (ekstern), Visa, Avgift
        public string Text { get; set; }
        public decimal OutAccount { get; set; }
        public decimal InAccount { get; set; }
        public decimal AccountChange { get; set; }

        // estimated fields based on content
        public AccountingTypeEnum AccountingType { get; set; }
        public DateTime ExternalPurchaseDate { get; set; }
        public decimal ExternalPurchaseAmount { get; set; }
        public string ExternalPurchaseCurrency { get; set; }
        public string ExternalPurchaseVendor { get; set; }
        public decimal ExternalPurchaseExchangeRate { get; set; }

        public override string ToString()
        {
            string tmpString = "";
            switch (this.AccountingType)
            {
                case AccountingTypeEnum.CostOfGoods:
                case AccountingTypeEnum.CostOfAdvertising:
                case AccountingTypeEnum.CostOfWebShop:
                case AccountingTypeEnum.CostOfDomain:
                case AccountingTypeEnum.CostOfServer:
                    tmpString = string.Format("{0:dd.MM.yyyy} {1} {2} {3:dd.MM} {4} {5} {6} {7:C}", TransactionDate, GetAccountingTypeString(), ArchiveReference, ExternalPurchaseDate, ExternalPurchaseVendor, ExternalPurchaseAmount, ExternalPurchaseCurrency, AccountChange);
                    break;
                case AccountingTypeEnum.TransferPaypal:
                case AccountingTypeEnum.TransferStripe:
                    tmpString = string.Format("{0:dd.MM.yyyy} {1} {2} {3:dd.MM} {4} {5:C}", TransactionDate, GetAccountingTypeString(), ArchiveReference, ExternalPurchaseDate, ExternalPurchaseVendor, AccountChange);
                    break;
                case AccountingTypeEnum.TransferUnknown:
                case AccountingTypeEnum.IncomeReturn:
                case AccountingTypeEnum.IncomeInterest:
                case AccountingTypeEnum.IncomeUnknown:
                case AccountingTypeEnum.CostOfBank:
                case AccountingTypeEnum.CostOfInvoice:
                case AccountingTypeEnum.CostOfTryouts:
                case AccountingTypeEnum.CostUnknown:
                    tmpString = string.Format("{0:dd.MM.yyyy} {1} {2} {3} {4} {5:C}", TransactionDate, GetAccountingTypeString(), ArchiveReference, Type, Text, AccountChange);
                    break;
            }

            return tmpString;
        }

        public string GetAccountingTypeString()
        {
            var accountingTypeString = "";
            switch (this.AccountingType)
            {
                case AccountingTypeEnum.CostOfGoods:
                    accountingTypeString = "KOST VARE";
                    break;
                case AccountingTypeEnum.CostOfAdvertising:
                    accountingTypeString = "KOST REKLAME";
                    break;
                case AccountingTypeEnum.CostOfWebShop:
                    accountingTypeString = "KOST NETTBUTIKK";
                    break;
                case AccountingTypeEnum.CostOfDomain:
                    accountingTypeString = "KOST DOMENE";
                    break;
                case AccountingTypeEnum.CostOfServer:
                    accountingTypeString = "KOST SERVER";
                    break;
                case AccountingTypeEnum.CostOfBank:
                    accountingTypeString = "KOST AVGIFT";
                    break;
                case AccountingTypeEnum.CostOfInvoice:
                    accountingTypeString = "KOST GIRO";
                    break;
                case AccountingTypeEnum.CostOfTryouts:
                    accountingTypeString = "KOST PRØVE";
                    break;
                case AccountingTypeEnum.CostUnknown:
                    accountingTypeString = "KOST UKJENT";
                    break;
                case AccountingTypeEnum.TransferPaypal:
                    accountingTypeString = "OVERFØRSEL PAYPAL";
                    break;
                case AccountingTypeEnum.TransferStripe:
                    accountingTypeString = "OVERFØRSEL STRIPE";
                    break;
                case AccountingTypeEnum.TransferUnknown:
                    accountingTypeString = "OVERFØRSEL UKJENT";
                    break;
                case AccountingTypeEnum.IncomeReturn:
                    accountingTypeString = "INNTEKT RETUR";
                    break;
                case AccountingTypeEnum.IncomeInterest:
                    accountingTypeString = "INNTEKT RENTER";
                    break;
                case AccountingTypeEnum.IncomeUnknown:
                    accountingTypeString = "INNTEKT UKJENT";
                    break;
            }

            return accountingTypeString;
        }

        private Tuple<int, int> ParseDayAndMonth(string dayAndMonth)
        {
            var splitted = dayAndMonth.Split('.');
            int day = int.Parse(splitted.First());
            int month = int.Parse(splitted.Last());
            return new Tuple<int, int>(day, month);
        }

        public void ExtractAccountingInformation()
        {
            // good regexp tester   
            // https://regex101.com/

            if (Type.Equals("Kreditrente"))
            {
                this.AccountingType = AccountingTypeEnum.IncomeInterest;
                return;
            }
            else if (Type.Equals("Avgift"))
            {
                this.AccountingType = AccountingTypeEnum.CostOfBank;
                return;
            }
            else if (Type.Equals("Giro m/KID"))
            {
                this.AccountingType = AccountingTypeEnum.CostOfInvoice;
                return;
            }

            // check if the text is a purchase
            var matchPurchase = purchasePattern.Match(Text);
            if (matchPurchase.Success)
            {
                var dayAndMonth = matchPurchase.Groups[2].ToString();
                var currency = matchPurchase.Groups[3].ToString();
                var amount = matchPurchase.Groups[4].ToString();
                var vendor = matchPurchase.Groups[5].ToString();
                var exchangeRate = matchPurchase.Groups[7].ToString();

                // fix edge case where year is likely last year 
                var dayAndMonthParsed = ParseDayAndMonth(dayAndMonth);
                int day = dayAndMonthParsed.Item1;
                int month = dayAndMonthParsed.Item2;
                int year = TransactionDate.Year;
                var cutoffDate = new DateTime(year, 1, 10); // 10th of Jan
                if (TransactionDate < cutoffDate
                    && month == 12) // December
                {
                    year--;
                }
                var dateString = string.Format("{0}.{1}", dayAndMonth, year);
                // parse date
                DateTime purchaseDate = DateTime.ParseExact(dateString, "dd.MM.yyyy", CultureInfo.InvariantCulture);

                // store properies
                ExternalPurchaseDate = purchaseDate;
                ExternalPurchaseAmount = ExcelUtils.GetDecimalFromExcelCurrencyString(amount);
                ExternalPurchaseCurrency = currency;
                ExternalPurchaseVendor = vendor;
                ExternalPurchaseExchangeRate = ExcelUtils.GetDecimalFromExcelCurrencyString(exchangeRate);

                if (vendor.CaseInsensitiveContains("Wazalo")
                    || vendor.CaseInsensitiveContains("Shopifycomc"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfWebShop;
                }
                else if (vendor.CaseInsensitiveContains("Facebk"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfAdvertising;
                }
                else if (vendor.CaseInsensitiveContains("Gandi"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfDomain;
                }
                else if (vendor.CaseInsensitiveContains("Scaleway"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfServer;
                }
                else if (vendor.CaseInsensitiveContains("AliExpress"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfGoods;
                }
                else
                {
                    this.AccountingType = AccountingTypeEnum.CostUnknown;
                }

                if (AccountChange >= 0)
                {
                    // not a purchase, but a return
                    this.AccountingType = AccountingTypeEnum.IncomeReturn;
                }
                return;
            }

            // if not a purchase, check if it is a transfer
            var matchTransfer = transferPattern.Match(Text);
            if (matchTransfer.Success)
            {
                var vendor = matchTransfer.Groups[1].Value.ToString();
                var date = matchTransfer.Groups[2].Value.ToString();

                // fix date
                DateTime purchaseDate = DateTime.ParseExact(date, "dd.MM.yy", CultureInfo.InvariantCulture);

                // store properties
                ExternalPurchaseDate = purchaseDate;
                ExternalPurchaseVendor = vendor;

                if (vendor.CaseInsensitiveContains("The Currency Cloud"))
                {
                    this.AccountingType = AccountingTypeEnum.TransferStripe;
                }
                else if (vendor.CaseInsensitiveContains("Paypal Pte Ltd"))
                {
                    this.AccountingType = AccountingTypeEnum.TransferPaypal;
                }
                else
                {
                    this.AccountingType = AccountingTypeEnum.TransferUnknown;
                }
                return;
            }

            // if neither match for purchase or transfer
            if (AccountChange > 0)
            {
                this.AccountingType = AccountingTypeEnum.IncomeUnknown;
            }
            else
            {
                if (Text.CaseInsensitiveContains("Gandi"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfDomain;
                }
                else if (Text.CaseInsensitiveContains("Prøvekjøp"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfTryouts;
                }
                else
                {
                    this.AccountingType = AccountingTypeEnum.CostUnknown;
                }
            }
        }

        public void ExtractAccountingInformationAPI()
        {
            if (Type.Equals("KREDITRTE"))
            {
                this.AccountingType = AccountingTypeEnum.IncomeInterest;
                return;
            }
            else if (Type.Equals("GEBYR"))
            {
                this.AccountingType = AccountingTypeEnum.CostOfBank;
                return;
            }
            else if (Type.Equals("NETTGIRO"))
            {
                this.AccountingType = AccountingTypeEnum.CostOfInvoice;
                return;
            }
            else if (Type.Equals("OVFNETTB"))
            {
                this.AccountingType = AccountingTypeEnum.TransferUnknown;
                return;
            }
            else if (Type.Equals("OVERFØRT"))
            {
                // check for transfer patterns
                var matchTransfer2 = transferPattern2.Match(Text);
                if (matchTransfer2.Success)
                {
                    var vendor = matchTransfer2.Groups[1].Value.ToString().Trim();
                    var currency = matchTransfer2.Groups[2].Value.ToString();
                    var amount = matchTransfer2.Groups[3].Value.ToString();

                    // store properties                    
                    bool isNorwegian = (amount.Contains(",") ? true : false); 
                    ExternalPurchaseAmount = ExcelUtils.GetDecimalFromExcelCurrencyString(amount, isNorwegian);
                    ExternalPurchaseCurrency = currency;
                    ExternalPurchaseVendor = vendor;
                    AccountingType = AccountingTypeEnum.TransferStripe;
                    ExternalPurchaseDate = TransactionDate;
                    return;
                }
            }

            // check if it is a purchase or return
            if (Type.Equals("VISA VARE"))
            {
                var vendor = ExternalPurchaseVendor;

                if (vendor.CaseInsensitiveContains("Wazalo")
                    || vendor.CaseInsensitiveContains("Shopifycomc")
                    || vendor.CaseInsensitiveContains("SHOPIFY.COM/C")
                    || vendor.CaseInsensitiveContains("SHOPIFY"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfWebShop;
                }
                else if (vendor.CaseInsensitiveContains("Facebk"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfAdvertising;
                }
                else if (vendor.CaseInsensitiveContains("Gandi"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfDomain;
                }
                else if (vendor.CaseInsensitiveContains("Scaleway"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfServer;
                }
                else if (vendor.CaseInsensitiveContains("AliExpress"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfGoods;
                }
                else if (vendor.CaseInsensitiveContains("PayPal"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfGoods;
                }
                else
                {
                    this.AccountingType = AccountingTypeEnum.CostUnknown;
                }

                if (AccountChange >= 0)
                {
                    // not a purchase, but a return
                    this.AccountingType = AccountingTypeEnum.IncomeReturn;
                }
                return;
            }

            // if not a purchase, check if it is a transfer
            var matchTransfer = transferPattern.Match(Text);
            if (matchTransfer.Success)
            {
                var vendor = matchTransfer.Groups[1].Value.ToString();
                var date = matchTransfer.Groups[2].Value.ToString();

                // fix date
                DateTime purchaseDate = DateTime.ParseExact(date, "dd.MM.yy", CultureInfo.InvariantCulture);

                // store properties
                ExternalPurchaseDate = purchaseDate;
                ExternalPurchaseVendor = vendor;

                if (vendor.CaseInsensitiveContains("The Currency Cloud"))
                {
                    this.AccountingType = AccountingTypeEnum.TransferStripe;
                }
                else if (vendor.CaseInsensitiveContains("Paypal Pte Ltd"))
                {
                    this.AccountingType = AccountingTypeEnum.TransferPaypal;
                }
                else
                {
                    this.AccountingType = AccountingTypeEnum.TransferUnknown;
                }
                return;
            }

            // if neither match for purchase or transfer
            if (AccountChange > 0)
            {
                this.AccountingType = AccountingTypeEnum.IncomeUnknown;
            }
            else
            {
                if (Text.CaseInsensitiveContains("Gandi"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfDomain;
                }
                else if (Text.CaseInsensitiveContains("Prøvekjøp"))
                {
                    this.AccountingType = AccountingTypeEnum.CostOfTryouts;
                }
                else
                {
                    this.AccountingType = AccountingTypeEnum.CostUnknown;
                }
            }
        }
    }
}
