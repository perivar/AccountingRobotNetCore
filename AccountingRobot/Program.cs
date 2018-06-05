using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using System.Linq;
using ClosedXML.Excel;
using System.Data;
using System.IO;
using System.Globalization;
using AccountingServices;
using System.Text.RegularExpressions;
using Google.Apis.Sheets.v4.Data;
using System.Data.Common;

namespace AccountingRobot
{
    partial class Program
    {
        const string GOOGLE_SHEET_NAME = "BILAGSJOURNAL2";

        static void Main(string[] args)
        {
            bool USE_EXCEL = false;

            // init date
            var date = new Date();

            IMyConfiguration configuration = new MyConfiguration();

            // prepopulate lookup lists
            Console.Out.WriteLine("Prepopulating Lookup Lists ...");

            var stripeTransactions = StripeChargeFactory.Instance.GetLatest(configuration, false);
            Console.Out.WriteLine("Successfully read Stripe transactions ...");

            var paypalTransactions = PayPalFactory.Instance.GetLatest(configuration, false);
            Console.Out.WriteLine("Successfully read PayPal transactions ...");

            // process the transactions and create accounting overview
            var customerNames = new List<string>();
            var accountingShopifyItems = ProcessShopifyStatement(configuration, customerNames, stripeTransactions, paypalTransactions);

            // select only distinct 
            customerNames = customerNames.Distinct().ToList();

            // find latest skandiabanken transaction spreadsheet
            var sBankenTransactions = SBankenFactory.Instance.GetLatest(configuration, true);
            var sBankenBankStatement = SBanken.GetBankStatementFromTransactions(sBankenTransactions);
            if (sBankenBankStatement.Transactions.Count() == 0)
            {
                // No transactions read, quitting
                Console.WriteLine("ERROR! No SBanken transactions read. Quitting!");
                Console.ReadLine();
                return;
            }
            var accountingBankItems = ProcessBankAccountStatement(configuration, sBankenBankStatement, customerNames, stripeTransactions, paypalTransactions);

            // merge into one list
            accountingShopifyItems.AddRange(accountingBankItems);

            // and sort (by ascending)
            var accountingItems = accountingShopifyItems.OrderBy(o => o.Date).ToList();

            // export or update accounting spreadsheet
            string accountingFileDir = configuration.GetValue("AccountingDir");
            string accountingFileNamePrefix = "wazalo regnskap";
            string accountingDateFromToRegexPattern = @"(\d{4}\-\d{2}\-\d{2})\-(\d{4}\-\d{2}\-\d{2})\.xlsx$";

            FileDate lastAccountingFileInfo = null;
            if (USE_EXCEL) lastAccountingFileInfo = Utils.FindLastCacheFile(accountingFileDir, accountingFileNamePrefix, accountingDateFromToRegexPattern, "yyyy-MM-dd", "\\-");

            // if the cache file object has values
            if (null != lastAccountingFileInfo && !lastAccountingFileInfo.Equals(default(FileDate)))
            {
                Console.Out.WriteLine("Found an accounting spreadsheet from {0:yyyy-MM-dd}", lastAccountingFileInfo.From);
                UpdateExcelFile(lastAccountingFileInfo.FilePath, accountingItems);

                // rename spreadsheet to today's date
                if (lastAccountingFileInfo.To != date.CurrentDate.Date)
                {
                    var from = lastAccountingFileInfo.From;
                    var to = date.CurrentDate;

                    string accountingFileName = string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.xlsx", accountingFileNamePrefix, from, to);
                    string filePath = Path.Combine(accountingFileDir, accountingFileName);

                    File.Move(lastAccountingFileInfo.FilePath, filePath);

                    Console.Out.WriteLine("Successfully renamed accounting file!");
                }
            }
            else
            {
                Console.Out.WriteLine("No existing accounting spreadsheets found - creating ...");

                var from = date.FirstDayOfTheYear;
                var to = date.CurrentDate;

                if (USE_EXCEL)
                {
                    // export to excel file
                    string accountingFileName = string.Format("{0}-{1:yyyy-MM-dd}-{2:yyyy-MM-dd}.xlsx", accountingFileNamePrefix, from, to);
                    string filePath = Path.Combine(accountingFileDir, accountingFileName);

                    ExportToExcel(filePath, accountingItems);
                }
                else
                {
                    using (var googleSheetsFactory = new GoogleSheetsFactory())
                    {
                        int sheetId = googleSheetsFactory.GetSheetIdFromSheetName(GOOGLE_SHEET_NAME);
                        if (sheetId > 0)
                        {
                            UpdateGoogleSheets(googleSheetsFactory, accountingItems);
                        }
                        else
                        {
                            ExportToGoogleSheets(googleSheetsFactory, accountingItems);
                        }
                    }
                }
            }

            Console.ReadLine();
        }

        #region Google Sheets Methods
        private static string[] GetAccountingHeaders()
        {
            // add accounting headers
            var accountingHeaders = new string[53];
            accountingHeaders[0] = "Næringsoppgave";
            accountingHeaders[16] = "1910";
            accountingHeaders[17] = "1912";
            accountingHeaders[18] = "1914";
            accountingHeaders[19] = "1920";

            accountingHeaders[22] = "2740";
            accountingHeaders[23] = "3000";
            accountingHeaders[24] = "3100";
            accountingHeaders[25] = "4005";
            accountingHeaders[26] = "4300";
            accountingHeaders[27] = "5000";
            accountingHeaders[28] = "5400";
            accountingHeaders[29] = "6000";
            accountingHeaders[30] = "6100";
            accountingHeaders[31] = "6340";
            accountingHeaders[32] = "6500";
            accountingHeaders[33] = "6695";
            accountingHeaders[34] = "6800";
            accountingHeaders[35] = "6810";
            accountingHeaders[36] = "6900";
            accountingHeaders[37] = "7098";
            accountingHeaders[38] = "7140";
            accountingHeaders[39] = "7330";
            accountingHeaders[40] = "7700";
            accountingHeaders[41] = "7770";
            accountingHeaders[42] = "7780";
            accountingHeaders[43] = "7785";
            accountingHeaders[44] = "7790";
            accountingHeaders[45] = "8099";
            accountingHeaders[46] = "8199";
            accountingHeaders[47] = "1200";
            accountingHeaders[48] = "1500";
            return accountingHeaders;
        }

        static void AppendDataTable(GoogleSheetsFactory googleSheetsFactory, List<AccountingItem> accountingItems, bool doAddSheet, bool doUseAccountingHeaders, bool doUseTableHeaders, bool doAutoResizeColumns, bool doUseSubTotals, bool doHideColumns, bool doAddBasicFilter)
        {
            var dt = GetDataTable(accountingItems);

            if (dt == null) return;

            // find or add google sheets spreadsheet 
            int sheetId = -1;
            if (doAddSheet)
            {
                sheetId = googleSheetsFactory.AddSheet(GOOGLE_SHEET_NAME, dt.Columns.Count);
            }
            else
            {
                sheetId = googleSheetsFactory.GetSheetIdFromSheetName(GOOGLE_SHEET_NAME);
            }

            int startColumnIndex = 0;
            int endColumnIndex = dt.Columns.Count;
            int startRowIndex = 0;
            int endRowIndex = dt.Rows.Count;

            using (var googleBatchUpdateRequest = new GoogleSheetsBatchUpdateRequests())
            {
                // append headers
                if (doUseAccountingHeaders)
                {
                    var accountingHeaders = GetAccountingHeaders();
                    googleBatchUpdateRequest.Add(GoogleSheetsRequests.GetAppendCellsRequest(sheetId, accountingHeaders, 0xFFFFFF, 0x000000));
                }

                // append data table in row under the headers, typically row 2
                googleBatchUpdateRequest.Add(GoogleSheetsRequests.GetAppendDataTableRequests(sheetId, dt, 0x000000, 0xFFFFFF, 0x000000, 0xdbe5f1, doUseTableHeaders));

                // auto resize columns
                if (doAutoResizeColumns)
                {
                    googleBatchUpdateRequest.Add(GoogleSheetsRequests.GetAutoResizeColumnsRequest(sheetId, startColumnIndex, endColumnIndex));
                }

                // hide archive reference and transaction id
                if (doHideColumns)
                {
                    googleBatchUpdateRequest.Add(
                        GoogleSheetsRequests.HideColumnsRequest(sheetId, "E", "F")
                    );
                }

                ApplyGoogleSheetFormatting(googleBatchUpdateRequest, sheetId, startRowIndex + 2, endRowIndex + 2, doUseSubTotals, doAddBasicFilter);

                googleBatchUpdateRequest.Execute();
            }
        }

        static void InsertDataTable(GoogleSheetsFactory googleSheetsFactory, DataTable dt, bool doAddSheet, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex, bool doUseSubTotals, bool doAddBasicFilter)
        {
            // find or add google sheets spreadsheet 
            int sheetId = -1;
            if (doAddSheet)
            {
                sheetId = googleSheetsFactory.AddSheet(GOOGLE_SHEET_NAME, dt.Columns.Count);
            }
            else
            {
                sheetId = googleSheetsFactory.GetSheetIdFromSheetName(GOOGLE_SHEET_NAME);
            }

            using (var googleBatchUpdateRequest = new GoogleSheetsBatchUpdateRequests())
            {
                googleBatchUpdateRequest.Add(
                    GoogleSheetsRequests.GetInsertDataTableRequests(sheetId,
                    dt,
                    startRowIndex, startColumnIndex,
                    0x000000, 0xFFFFFF, 0x000000, 0xdbe5f1, false));

                ApplyGoogleSheetFormatting(googleBatchUpdateRequest, sheetId, startRowIndex, endRowIndex, doUseSubTotals, doAddBasicFilter);

                googleBatchUpdateRequest.Execute();
            }
        }

        static void ApplyGoogleSheetFormatting(GoogleSheetsBatchUpdateRequests googleBatchUpdateRequest, int sheetId, int startRowIndex, int endRowIndex, bool doUseSubTotals, bool doAddBasicFilter)
        {
            // calculate row offset for the forumlas
            int formulaRowIndex = startRowIndex + 1;
            int subTotalStartRow = 3;
            int periodStartRowIndex = 2;
            int basicFilterStartRowIndex = 1;

            // insert subtotal in last row
            if (doUseSubTotals)
            {
                // =SUBTOTAL(109;O3:O174) = sum and ignore hidden values
                googleBatchUpdateRequest.Add(
                    GoogleSheetsRequests.GetFormulaAndNumberFormatRequest(sheetId,
                    string.Format("=SUBTOTAL(109;Q{0}:Q{1})", subTotalStartRow, endRowIndex),
                    endRowIndex, endRowIndex + 1, "Q", "AY")
                );
            }

            // set int format to Periode and Bilagsnr
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetNumberFormatRequest(sheetId,
                "0", 0x000000, 0xdbe5f1, periodStartRowIndex, endRowIndex, "B", "B"));

            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetNumberFormatRequest(sheetId,
                "0", 0x000000, 0xdbe5f1, periodStartRowIndex, endRowIndex, "D", "D"));


            // set basic filter for all rows, can only be applied to header
            if (doAddBasicFilter)
            {
                googleBatchUpdateRequest.Add(
                    GoogleSheetsRequests.GetBasicFilterRequest(sheetId,
                    basicFilterStartRowIndex, endRowIndex, "A", "BA"));
            }

            // insert control formula in column 1
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetFormulaAndTextFormatRequest(sheetId,
                string.Format("=IF(BA{0}=0;\" \";\"!!FEIL!!\")", formulaRowIndex),
                0xFF0000, 0xEAF1FA,
                startRowIndex, endRowIndex, 0, 1)
            );

            // insert sum pre rounding formula in next last column 
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetFormulaAndNumberFormatRequest(sheetId,
                string.Format("=SUM(Q{0}:AY{0})", formulaRowIndex),
                startRowIndex, endRowIndex, "AZ", "AZ")
            );

            // insert sum rounding formula in last column
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetFormulaAndNumberFormatRequest(sheetId,
                string.Format("=ROUND(AZ{0};2)", formulaRowIndex),
                startRowIndex, endRowIndex, "BA", "BA")
            );

            // add VAT Sales column
            string vatSales = string.Format("=IF(AND(P{0}=\"NOK\";H{0}=\"SHOPIFY\");-(O{0}/1,25)*0,25;\" \")", formulaRowIndex);
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetFormulaRequest(sheetId,
                vatSales,
                startRowIndex, endRowIndex, "V", "V")
            );

            // add VAT Exempt column
            string salesVATExempt = string.Format("=IF(AND(P{0}=\"NOK\";H{0}=\"SHOPIFY\");-(O{0}/1,25);\" \")", formulaRowIndex);
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetFormulaRequest(sheetId,
                salesVATExempt,
                startRowIndex, endRowIndex, "X", "X")
            );

            // set colors green
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetFormatRequest(sheetId,
                0x000000, 0xEBF1DE,
                startRowIndex, endRowIndex, "U", "V")
            );

            // set colors blue
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetFormatRequest(sheetId,
                0x000000, 0xC5D9F1,
                startRowIndex, endRowIndex, "AV", "AY")
            );

            // set colors red
            googleBatchUpdateRequest.Add(
                GoogleSheetsRequests.GetFormatRequest(sheetId,
                0x000000, 0xF2DCDB,
                startRowIndex, endRowIndex, "AZ", "BA")
            );
        }

        static void ExportToGoogleSheets(GoogleSheetsFactory googleSheetsFactory, List<AccountingItem> accountingItems)
        {
            AppendDataTable(googleSheetsFactory, accountingItems, true, true, true, true, true, true, true);
            Console.Out.WriteLine("Successfully wrote accounting file to Google Sheets");
        }

        static void UpdateGoogleSheets(GoogleSheetsFactory googleSheetsFactory, List<AccountingItem> newAccountingItems)
        {
            // Update Google Sheets spreadsheet 
            int sheetId = googleSheetsFactory.GetSheetIdFromSheetName(GOOGLE_SHEET_NAME);

            var dt = googleSheetsFactory.ReadDataTable(GOOGLE_SHEET_NAME, 2);

            var existingAccountingItems = new Dictionary<DataRow, AccountingItem>();
            foreach (DataRow row in dt.Rows)
            {
                var accountingItem = new AccountingItem();
                // cSpell:disable
                accountingItem.Date = GoogleSheetsUtils.GetField<DateTime>(row, "Dato");
                accountingItem.Number = GoogleSheetsUtils.GetField<int>(row, "Bilagsnr.");
                accountingItem.ArchiveReference = GoogleSheetsUtils.GetField<string>(row, "Arkivreferanse");
                accountingItem.TransactionID = GoogleSheetsUtils.GetField<string>(row, "TransaksjonsId");
                accountingItem.Type = GoogleSheetsUtils.GetField<string>(row, "Type");
                accountingItem.AccountingType = GoogleSheetsUtils.GetField<string>(row, "Regnskapstype");
                accountingItem.Text = GoogleSheetsUtils.GetField<string>(row, "Tekst");
                accountingItem.CustomerName = GoogleSheetsUtils.GetField<string>(row, "Kundenavn");
                accountingItem.ErrorMessage = GoogleSheetsUtils.GetField<string>(row, "Feilmelding");
                accountingItem.Gateway = GoogleSheetsUtils.GetField<string>(row, "Gateway");
                accountingItem.NumSale = GoogleSheetsUtils.GetField<string>(row, "Num Salg");
                accountingItem.NumPurchase = GoogleSheetsUtils.GetField<string>(row, "Num Kjøp");
                accountingItem.PurchaseOtherCurrency = GoogleSheetsUtils.GetField<decimal>(row, "Kjøp annen valuta");
                accountingItem.OtherCurrency = GoogleSheetsUtils.GetField<string>(row, "Annen valuta");

                accountingItem.AccountPaypal = GoogleSheetsUtils.GetField<decimal>(row, "Paypal");    // 1910
                accountingItem.AccountStripe = GoogleSheetsUtils.GetField<decimal>(row, "Stripe");    // 1915
                accountingItem.AccountVipps = GoogleSheetsUtils.GetField<decimal>(row, "Vipps");  // 1918
                accountingItem.AccountBank = GoogleSheetsUtils.GetField<decimal>(row, "Bank");    // 1920

                accountingItem.VATPurchase = GoogleSheetsUtils.GetField<decimal>(row, "MVA Kjøp");
                accountingItem.VATSales = GoogleSheetsUtils.GetField<decimal>(row, "MVA Salg");

                accountingItem.VATSettlementAccount = GoogleSheetsUtils.GetField<decimal>(row, "Oppgjørskonto mva");
                accountingItem.SalesVAT = GoogleSheetsUtils.GetField<decimal>(row, "Salg mva-pliktig");   // 3000
                accountingItem.SalesVATExempt = GoogleSheetsUtils.GetField<decimal>(row, "Salg avgiftsfritt");    // 3100

                accountingItem.CostOfGoods = GoogleSheetsUtils.GetField<decimal>(row, "Varekostnad"); // 4005
                accountingItem.CostForReselling = GoogleSheetsUtils.GetField<decimal>(row, "Forbruk for videresalg"); // 4300
                accountingItem.CostForSalary = GoogleSheetsUtils.GetField<decimal>(row, "Lønn");  // 5000
                accountingItem.CostForSalaryTax = GoogleSheetsUtils.GetField<decimal>(row, "Arb.giver avgift");   // 5400
                accountingItem.CostForDepreciation = GoogleSheetsUtils.GetField<decimal>(row, "Avskrivninger");   // 6000
                accountingItem.CostForShipping = GoogleSheetsUtils.GetField<decimal>(row, "Frakt");   // 6100
                accountingItem.CostForElectricity = GoogleSheetsUtils.GetField<decimal>(row, "Strøm");    // 6340 
                accountingItem.CostForToolsInventory = GoogleSheetsUtils.GetField<decimal>(row, "Verktøy inventar");  // 6500
                accountingItem.CostForMaintenance = GoogleSheetsUtils.GetField<decimal>(row, "Vedlikehold");  // 6695
                accountingItem.CostForFacilities = GoogleSheetsUtils.GetField<decimal>(row, "Kontorkostnader");   // 6800 

                accountingItem.CostOfData = GoogleSheetsUtils.GetField<decimal>(row, "Datakostnader");    // 6810 
                accountingItem.CostOfPhoneInternetUse = GoogleSheetsUtils.GetField<decimal>(row, "Telefon Internett Bruk");   // 6900
                accountingItem.PrivateUseOfECom = GoogleSheetsUtils.GetField<decimal>(row, "Privat bruk av el.kommunikasjon");    // 7098
                accountingItem.CostForTravelAndAllowance = GoogleSheetsUtils.GetField<decimal>(row, "Reise og Diett");    // 7140
                accountingItem.CostOfAdvertising = GoogleSheetsUtils.GetField<decimal>(row, "Reklamekostnader");  // 7330
                accountingItem.CostOfOther = GoogleSheetsUtils.GetField<decimal>(row, "Diverse annet");   // 7700

                accountingItem.FeesBank = GoogleSheetsUtils.GetField<decimal>(row, "Gebyrer Bank");   // 7770
                accountingItem.FeesPaypal = GoogleSheetsUtils.GetField<decimal>(row, "Gebyrer Paypal");   // 7780
                accountingItem.FeesStripe = GoogleSheetsUtils.GetField<decimal>(row, "Gebyrer Stripe");   // 7785 

                accountingItem.CostForEstablishment = GoogleSheetsUtils.GetField<decimal>(row, "Etableringskostnader");   // 7790

                accountingItem.IncomeFinance = GoogleSheetsUtils.GetField<decimal>(row, "Finansinntekter");   // 8099
                accountingItem.CostOfFinance = GoogleSheetsUtils.GetField<decimal>(row, "Finanskostnader");   // 8199

                accountingItem.Investments = GoogleSheetsUtils.GetField<decimal>(row, "Investeringer");   // 1200
                accountingItem.AccountsReceivable = GoogleSheetsUtils.GetField<decimal>(row, "Kundefordringer");  // 1500
                accountingItem.PersonalWithdrawal = GoogleSheetsUtils.GetField<decimal>(row, "Privat uttak");
                accountingItem.PersonalDeposit = GoogleSheetsUtils.GetField<decimal>(row, "Privat innskudd");
                // cSpell:enable

                existingAccountingItems.Add(row, accountingItem);
            }

            // reduce the old Accounting Spreadsheet and remove the entries that doesn't have a number
            var existingAccountingItemsToDelete =
                (from row in existingAccountingItems
                 where
                 row.Value.Number == 0
                 orderby row.Value.Number ascending
                 select row);

            // identify elements from the new accounting items list that does not exist in the existing spreadsheet
            var existingAccountingItemsToKeep = existingAccountingItems.Except(existingAccountingItemsToDelete);
            var newAccountingElements = newAccountingItems.Except(existingAccountingItemsToKeep.Select(o => o.Value)).ToList();

            // delete rows from table
            int deleteRowTotalCount = existingAccountingItemsToDelete.Count();
            Console.Out.WriteLine("Deleting {0} rows", deleteRowTotalCount);
            if (deleteRowTotalCount > 0)
            {
                using (var googleBatchDeleteRequest = new GoogleSheetsBatchUpdateRequests())
                {
                    int startRowNumber = existingAccountingItemsToDelete.FirstOrDefault().Key.Field<int>("RowNumber");
                    int endRowNumber = existingAccountingItemsToDelete.Last().Key.Field<int>("RowNumber");
                    googleBatchDeleteRequest.Add(GoogleSheetsRequests.GetDeleteRowsRequest(sheetId, startRowNumber - 1, endRowNumber + 1)); // delete sub totals as well
                    googleBatchDeleteRequest.Execute();
                }
            }

            // how many new rows needs to be added
            int newRowTotalCount = newAccountingElements.Count();
            int existingTotalCount = existingAccountingItemsToKeep.Count();
            Console.Out.WriteLine("\nAppending {0} rows", newRowTotalCount);
            if (newRowTotalCount > 0)
            {
                int startRowNumber = 0;
                using (var googleBatchInsertRequest = new GoogleSheetsBatchUpdateRequests())
                {
                    if (existingTotalCount > 0)
                    {
                        startRowNumber = existingAccountingItemsToKeep.Last().Key.Field<int>("RowNumber");
                    }
                    else
                    {
                        startRowNumber = 2;
                    }
                    googleBatchInsertRequest.Add(GoogleSheetsRequests.GetInsertRowsRequest(sheetId, startRowNumber, startRowNumber + newRowTotalCount, true));
                    googleBatchInsertRequest.Execute();
                }

                var dtToInsert = GetDataTable(newAccountingElements);
                if (dtToInsert == null) return;

                int startRowIndex = startRowNumber;
                int endRowIndex = startRowNumber + newRowTotalCount;
                int startColumnIndex = 0;
                int endColumnIndex = dtToInsert.Columns.Count + 1;
                InsertDataTable(googleSheetsFactory, dtToInsert, false, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex, true, true);

                Console.Out.WriteLine("Successfully updated accounting file!");
            }
        }
        #endregion

        #region Excel Methods
        static void ExportToExcel(string filePath, List<AccountingItem> accountingItems)
        {
            var dt = GetDataTable(accountingItems);

            // Build Excel spreadsheet using Closed XML
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Bilagsjournal");

                // add accounting headers
                ws.Cell(1, 1).Value = "Næringsoppgave";
                ws.Cell(1, 17).Value = "1910";
                ws.Cell(1, 18).Value = "1912";
                ws.Cell(1, 19).Value = "1914";
                ws.Cell(1, 20).Value = "1920";

                ws.Cell(1, 23).Value = "2740";
                ws.Cell(1, 24).Value = "3000";
                ws.Cell(1, 25).Value = "3100";
                ws.Cell(1, 26).Value = "4005";
                ws.Cell(1, 27).Value = "4300";
                ws.Cell(1, 28).Value = "5000";
                ws.Cell(1, 29).Value = "5400";
                ws.Cell(1, 30).Value = "6000";
                ws.Cell(1, 31).Value = "6100";
                ws.Cell(1, 32).Value = "6340";
                ws.Cell(1, 33).Value = "6500";
                ws.Cell(1, 34).Value = "6695";
                ws.Cell(1, 35).Value = "6800";
                ws.Cell(1, 36).Value = "6810";
                ws.Cell(1, 37).Value = "6900";
                ws.Cell(1, 38).Value = "7098";
                ws.Cell(1, 39).Value = "7140";
                ws.Cell(1, 40).Value = "7330";
                ws.Cell(1, 41).Value = "7700";
                ws.Cell(1, 42).Value = "7770";
                ws.Cell(1, 43).Value = "7780";
                ws.Cell(1, 44).Value = "7785";
                ws.Cell(1, 45).Value = "7790";
                ws.Cell(1, 46).Value = "8099";
                ws.Cell(1, 47).Value = "8199";
                ws.Cell(1, 48).Value = "1200";
                ws.Cell(1, 49).Value = "1500";

                // set font color for header range
                var headerRange = ws.Range("A1:BA1");
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Fill.BackgroundColor = XLColor.Black;

                // insert datatable in row 2
                var table = ws.Cell(2, 1).InsertTable(dt);

                table.Theme = XLTableTheme.TableStyleLight16;

                // turn on table total rows and set the functions for each of the relevant columns
                SetExcelTableTotalsRowFunction(table);

                if (table != null)
                {
                    foreach (var row in table.DataRange.Rows())
                    {
                        SetExcelRowFormulas(row);
                        SetExcelRowStyles(row);
                    }
                }

                // resize
                //ws.Columns().AdjustToContents();  // Adjust column width
                //ws.Rows().AdjustToContents();     // Adjust row heights

                wb.SaveAs(filePath);
                Console.Out.WriteLine("Successfully wrote accounting file to {0}", filePath);
            }
        }

        static void UpdateExcelFile(string filePath, List<AccountingItem> newAccountingItems)
        {
            // go through each row and check if it has already been "fixed".
            // i.e. the Number columns is no longer 0

            XLWorkbook wb = new XLWorkbook(filePath);
            IXLWorksheet ws = wb.Worksheet("Bilagsjournal");

            IXLTables tables = ws.Tables;
            IXLTable table = tables.FirstOrDefault();

            var existingAccountingItems = new Dictionary<IXLTableRow, AccountingItem>();
            if (table != null)
            {
                foreach (var row in table.DataRange.Rows())
                {
                    var accountingItem = new AccountingItem();
                    // cSpell:disable
                    accountingItem.Date = ExcelUtils.GetField<DateTime>(row, "Dato");
                    accountingItem.Number = ExcelUtils.GetField<int>(row, "Bilagsnr.");
                    accountingItem.ArchiveReference = ExcelUtils.GetField<string>(row, "Arkivreferanse");
                    accountingItem.TransactionID = ExcelUtils.GetField<string>(row, "TransaksjonsId");
                    accountingItem.Type = ExcelUtils.GetField<string>(row, "Type");
                    accountingItem.AccountingType = ExcelUtils.GetField<string>(row, "Regnskapstype");
                    accountingItem.Text = ExcelUtils.GetField<string>(row, "Tekst");
                    accountingItem.CustomerName = ExcelUtils.GetField<string>(row, "Kundenavn");
                    accountingItem.ErrorMessage = ExcelUtils.GetField<string>(row, "Feilmelding");
                    accountingItem.Gateway = ExcelUtils.GetField<string>(row, "Gateway");
                    accountingItem.NumSale = ExcelUtils.GetField<string>(row, "Num Salg");
                    accountingItem.NumPurchase = ExcelUtils.GetField<string>(row, "Num Kjøp");
                    accountingItem.PurchaseOtherCurrency = ExcelUtils.GetField<decimal>(row, "Kjøp annen valuta");
                    accountingItem.OtherCurrency = ExcelUtils.GetField<string>(row, "Annen valuta");

                    accountingItem.AccountPaypal = ExcelUtils.GetField<decimal>(row, "Paypal");	// 1910
                    accountingItem.AccountStripe = ExcelUtils.GetField<decimal>(row, "Stripe");	// 1915
                    accountingItem.AccountVipps = ExcelUtils.GetField<decimal>(row, "Vipps");	// 1918
                    accountingItem.AccountBank = ExcelUtils.GetField<decimal>(row, "Bank");	// 1920

                    accountingItem.VATPurchase = ExcelUtils.GetField<decimal>(row, "MVA Kjøp");
                    accountingItem.VATSales = ExcelUtils.GetField<decimal>(row, "MVA Salg");

                    accountingItem.VATSettlementAccount = ExcelUtils.GetField<decimal>(row, "Oppgjørskonto mva");
                    accountingItem.SalesVAT = ExcelUtils.GetField<decimal>(row, "Salg mva-pliktig");	// 3000
                    accountingItem.SalesVATExempt = ExcelUtils.GetField<decimal>(row, "Salg avgiftsfritt");	// 3100

                    accountingItem.CostOfGoods = ExcelUtils.GetField<decimal>(row, "Varekostnad");	// 4005
                    accountingItem.CostForReselling = ExcelUtils.GetField<decimal>(row, "Forbruk for videresalg");	// 4300
                    accountingItem.CostForSalary = ExcelUtils.GetField<decimal>(row, "Lønn");	// 5000
                    accountingItem.CostForSalaryTax = ExcelUtils.GetField<decimal>(row, "Arb.giver avgift");	// 5400
                    accountingItem.CostForDepreciation = ExcelUtils.GetField<decimal>(row, "Avskrivninger");	// 6000
                    accountingItem.CostForShipping = ExcelUtils.GetField<decimal>(row, "Frakt");	// 6100
                    accountingItem.CostForElectricity = ExcelUtils.GetField<decimal>(row, "Strøm");	// 6340 
                    accountingItem.CostForToolsInventory = ExcelUtils.GetField<decimal>(row, "Verktøy inventar");	// 6500
                    accountingItem.CostForMaintenance = ExcelUtils.GetField<decimal>(row, "Vedlikehold");	// 6695
                    accountingItem.CostForFacilities = ExcelUtils.GetField<decimal>(row, "Kontorkostnader");	// 6800 

                    accountingItem.CostOfData = ExcelUtils.GetField<decimal>(row, "Datakostnader");	// 6810 
                    accountingItem.CostOfPhoneInternetUse = ExcelUtils.GetField<decimal>(row, "Telefon Internett Bruk");	// 6900
                    accountingItem.PrivateUseOfECom = ExcelUtils.GetField<decimal>(row, "Privat bruk av el.kommunikasjon");	// 7098
                    accountingItem.CostForTravelAndAllowance = ExcelUtils.GetField<decimal>(row, "Reise og Diett");	// 7140
                    accountingItem.CostOfAdvertising = ExcelUtils.GetField<decimal>(row, "Reklamekostnader");	// 7330
                    accountingItem.CostOfOther = ExcelUtils.GetField<decimal>(row, "Diverse annet");	// 7700

                    accountingItem.FeesBank = ExcelUtils.GetField<decimal>(row, "Gebyrer Bank");	// 7770
                    accountingItem.FeesPaypal = ExcelUtils.GetField<decimal>(row, "Gebyrer Paypal");	// 7780
                    accountingItem.FeesStripe = ExcelUtils.GetField<decimal>(row, "Gebyrer Stripe");	// 7785 

                    accountingItem.CostForEstablishment = ExcelUtils.GetField<decimal>(row, "Etableringskostnader");	// 7790

                    accountingItem.IncomeFinance = ExcelUtils.GetField<decimal>(row, "Finansinntekter");	// 8099
                    accountingItem.CostOfFinance = ExcelUtils.GetField<decimal>(row, "Finanskostnader");	// 8199

                    accountingItem.Investments = ExcelUtils.GetField<decimal>(row, "Investeringer");	// 1200
                    accountingItem.AccountsReceivable = ExcelUtils.GetField<decimal>(row, "Kundefordringer");	// 1500
                    accountingItem.PersonalWithdrawal = ExcelUtils.GetField<decimal>(row, "Privat uttak");
                    accountingItem.PersonalDeposit = ExcelUtils.GetField<decimal>(row, "Privat innskudd");
                    // cSpell:enable

                    existingAccountingItems.Add(row, accountingItem);
                }

                // reduce the old Accounting Spreadsheet and remove the entries that doesn't have a number
                var existingAccountingItemsToDelete =
                    (from row in existingAccountingItems
                     where
                     row.Value.Number == 0
                     orderby row.Value.Number ascending
                     select row);

                // identify elements from the new accounting items list that does not exist in the existing spreadsheet
                var existingAccountingItemsToKeep = existingAccountingItems.Except(existingAccountingItemsToDelete);
                var newAccountingElements = newAccountingItems.Except(existingAccountingItemsToKeep.Select(o => o.Value)).ToList();

                // delete rows from table
                int deleteRowCounter = 0;
                int deleteRowTotalCount = existingAccountingItemsToDelete.Count();
                Console.Out.WriteLine("Deleting {0} rows", deleteRowTotalCount);
                foreach (var deleteRow in existingAccountingItemsToDelete)
                {
                    deleteRowCounter++;
                    Console.Out.Write("\rDeleting row {0}/{1} ({2})", deleteRowCounter, deleteRowTotalCount, deleteRow.Key.RangeAddress);
                    deleteRow.Key.Delete(XLShiftDeletedCells.ShiftCellsUp);
                }

                // how many new rows needs to be added
                int newRowTotalCount = newAccountingElements.Count();
                Console.Out.WriteLine("\nAppending {0} rows", newRowTotalCount);
                if (newRowTotalCount > 0)
                {
                    // turn off totals row before adding more rows
                    table.ShowTotalsRow = false;

                    // insert new rows below the existing table
                    var newRows = table.InsertRowsBelow(newRowTotalCount, true);

                    // insert values in the new rows
                    var newRowCounter = 0;
                    foreach (var newRow in newRows)
                    {
                        newRow.Cell(1).Value = "";
                        newRow.Cell(2).Value = newAccountingElements[newRowCounter].Periode;
                        newRow.Cell(3).Value = newAccountingElements[newRowCounter].Date;
                        newRow.Cell(4).Value = newAccountingElements[newRowCounter].Number;
                        newRow.Cell(5).Value = newAccountingElements[newRowCounter].ArchiveReference;
                        newRow.Cell(6).Value = newAccountingElements[newRowCounter].TransactionID;
                        newRow.Cell(7).Value = newAccountingElements[newRowCounter].Type;
                        newRow.Cell(8).Value = newAccountingElements[newRowCounter].AccountingType;
                        newRow.Cell(9).Value = newAccountingElements[newRowCounter].Text;
                        newRow.Cell(10).Value = newAccountingElements[newRowCounter].CustomerName;
                        newRow.Cell(11).Value = newAccountingElements[newRowCounter].ErrorMessage;
                        newRow.Cell(12).Value = newAccountingElements[newRowCounter].Gateway;
                        newRow.Cell(13).Value = newAccountingElements[newRowCounter].NumSale;
                        newRow.Cell(14).Value = newAccountingElements[newRowCounter].NumPurchase;
                        newRow.Cell(15).Value = newAccountingElements[newRowCounter].PurchaseOtherCurrency;
                        newRow.Cell(16).Value = newAccountingElements[newRowCounter].OtherCurrency;

                        newRow.Cell(17).Value = newAccountingElements[newRowCounter].AccountPaypal;               // 1910
                        newRow.Cell(18).Value = newAccountingElements[newRowCounter].AccountStripe;               // 1912
                        newRow.Cell(19).Value = newAccountingElements[newRowCounter].AccountVipps;                // 1914
                        newRow.Cell(20).Value = newAccountingElements[newRowCounter].AccountBank;                 // 1920

                        newRow.Cell(21).Value = newAccountingElements[newRowCounter].VATPurchase;
                        newRow.Cell(22).Value = newAccountingElements[newRowCounter].VATSales;

                        newRow.Cell(23).Value = newAccountingElements[newRowCounter].VATSettlementAccount;        // 2740                        
                        newRow.Cell(24).Value = newAccountingElements[newRowCounter].SalesVAT;                    // 3000
                        newRow.Cell(25).Value = newAccountingElements[newRowCounter].SalesVATExempt;              // 3100

                        newRow.Cell(26).Value = newAccountingElements[newRowCounter].CostOfGoods;                 // 4005
                        newRow.Cell(27).Value = newAccountingElements[newRowCounter].CostForReselling;            // 4300
                        newRow.Cell(28).Value = newAccountingElements[newRowCounter].CostForSalary;               // 5000
                        newRow.Cell(29).Value = newAccountingElements[newRowCounter].CostForSalaryTax;            // 5400
                        newRow.Cell(30).Value = newAccountingElements[newRowCounter].CostForDepreciation;         // 6000
                        newRow.Cell(31).Value = newAccountingElements[newRowCounter].CostForShipping;             // 6100
                        newRow.Cell(32).Value = newAccountingElements[newRowCounter].CostForElectricity;          // 6340 
                        newRow.Cell(33).Value = newAccountingElements[newRowCounter].CostForToolsInventory;       // 6500
                        newRow.Cell(34).Value = newAccountingElements[newRowCounter].CostForMaintenance;          // 6695
                        newRow.Cell(35).Value = newAccountingElements[newRowCounter].CostForFacilities;           // 6800 

                        newRow.Cell(36).Value = newAccountingElements[newRowCounter].CostOfData;                  // 6810 
                        newRow.Cell(37).Value = newAccountingElements[newRowCounter].CostOfPhoneInternetUse;      // 6900
                        newRow.Cell(38).Value = newAccountingElements[newRowCounter].PrivateUseOfECom;            // 7098
                        newRow.Cell(39).Value = newAccountingElements[newRowCounter].CostForTravelAndAllowance;   // 7140
                        newRow.Cell(40).Value = newAccountingElements[newRowCounter].CostOfAdvertising;           // 7330
                        newRow.Cell(41).Value = newAccountingElements[newRowCounter].CostOfOther;                 // 7700

                        newRow.Cell(42).Value = newAccountingElements[newRowCounter].FeesBank;                    // 7770
                        newRow.Cell(43).Value = newAccountingElements[newRowCounter].FeesPaypal;                  // 7780
                        newRow.Cell(44).Value = newAccountingElements[newRowCounter].FeesStripe;                  // 7785 

                        newRow.Cell(45).Value = newAccountingElements[newRowCounter].CostForEstablishment;        // 7790

                        newRow.Cell(46).Value = newAccountingElements[newRowCounter].IncomeFinance;               // 8099
                        newRow.Cell(47).Value = newAccountingElements[newRowCounter].CostOfFinance;               // 8199

                        newRow.Cell(48).Value = newAccountingElements[newRowCounter].Investments;                 // 1200
                        newRow.Cell(49).Value = newAccountingElements[newRowCounter].AccountsReceivable;          // 1500
                        newRow.Cell(50).Value = newAccountingElements[newRowCounter].PersonalWithdrawal;
                        newRow.Cell(51).Value = newAccountingElements[newRowCounter].PersonalDeposit;

                        SetExcelRowFormulas(newRow);
                        SetExcelRowStyles(newRow);

                        newRowCounter++;
                    }

                    // turn on table total rows and set the functions for each of the relevant columns
                    SetExcelTableTotalsRowFunction(table);
                }
                else
                {
                    Console.Out.WriteLine("Nothing to update! Quitting.");
                    return;
                }
            }

            // resize
            //ws.Columns().AdjustToContents();  // Adjust column width
            //ws.Rows().AdjustToContents();     // Adjust row heights

            wb.Save();
            Console.Out.WriteLine("Successfully updated accounting file!");
        }

        static void UpdateExcelFileWithTransactionsIds(string filePath, List<AccountingItem> newAccountingItems)
        {
            XLWorkbook wb = new XLWorkbook(filePath);
            IXLWorksheet ws = wb.Worksheet("Bilagsjournal");

            IXLTables tables = ws.Tables;
            IXLTable table = tables.FirstOrDefault();

            var existingAccountingItems = new Dictionary<IXLTableRow, AccountingItem>();
            if (table != null)
            {
                foreach (var row in table.DataRange.Rows())
                {
                    var accountingItem = new AccountingItem();
                    // cSpell:disable
                    accountingItem.Date = ExcelUtils.GetField<DateTime>(row, "Dato");
                    accountingItem.Number = ExcelUtils.GetField<int>(row, "Bilagsnr.");
                    accountingItem.ArchiveReference = ExcelUtils.GetField<string>(row, "Arkivreferanse");
                    accountingItem.TransactionID = ExcelUtils.GetField<string>(row, "TransaksjonsId");
                    accountingItem.Type = ExcelUtils.GetField<string>(row, "Type");
                    accountingItem.AccountingType = ExcelUtils.GetField<string>(row, "Regnskapstype");
                    accountingItem.Text = ExcelUtils.GetField<string>(row, "Tekst");
                    accountingItem.CustomerName = ExcelUtils.GetField<string>(row, "Kundenavn");
                    accountingItem.ErrorMessage = ExcelUtils.GetField<string>(row, "Feilmelding");
                    accountingItem.Gateway = ExcelUtils.GetField<string>(row, "Gateway");
                    accountingItem.NumSale = ExcelUtils.GetField<string>(row, "Num Salg");
                    accountingItem.NumPurchase = ExcelUtils.GetField<string>(row, "Num Kjøp");
                    accountingItem.PurchaseOtherCurrency = ExcelUtils.GetField<decimal>(row, "Kjøp annen valuta");
                    accountingItem.OtherCurrency = ExcelUtils.GetField<string>(row, "Annen valuta");

                    accountingItem.AccountPaypal = ExcelUtils.GetField<decimal>(row, "Paypal");	// 1910
                    accountingItem.AccountStripe = ExcelUtils.GetField<decimal>(row, "Stripe");	// 1915
                    accountingItem.AccountVipps = ExcelUtils.GetField<decimal>(row, "Vipps");	// 1918
                    accountingItem.AccountBank = ExcelUtils.GetField<decimal>(row, "Bank");	// 1920

                    accountingItem.VATPurchase = ExcelUtils.GetField<decimal>(row, "MVA Kjøp");
                    accountingItem.VATSales = ExcelUtils.GetField<decimal>(row, "MVA Salg");

                    accountingItem.VATSettlementAccount = ExcelUtils.GetField<decimal>(row, "Oppgjørskonto mva"); // 2740
                    accountingItem.SalesVAT = ExcelUtils.GetField<decimal>(row, "Salg mva-pliktig");	// 3000
                    accountingItem.SalesVATExempt = ExcelUtils.GetField<decimal>(row, "Salg avgiftsfritt");	// 3100

                    accountingItem.CostOfGoods = ExcelUtils.GetField<decimal>(row, "Varekostnad");	// 4005
                    accountingItem.CostForReselling = ExcelUtils.GetField<decimal>(row, "Forbruk for videresalg");	// 4300
                    accountingItem.CostForSalary = ExcelUtils.GetField<decimal>(row, "Lønn");	// 5000
                    accountingItem.CostForSalaryTax = ExcelUtils.GetField<decimal>(row, "Arb.giver avgift");	// 5400
                    accountingItem.CostForDepreciation = ExcelUtils.GetField<decimal>(row, "Avskrivninger");	// 6000
                    accountingItem.CostForShipping = ExcelUtils.GetField<decimal>(row, "Frakt");	// 6100
                    accountingItem.CostForElectricity = ExcelUtils.GetField<decimal>(row, "Strøm");	// 6340 
                    accountingItem.CostForToolsInventory = ExcelUtils.GetField<decimal>(row, "Verktøy inventar");	// 6500
                    accountingItem.CostForMaintenance = ExcelUtils.GetField<decimal>(row, "Vedlikehold");	// 6695
                    accountingItem.CostForFacilities = ExcelUtils.GetField<decimal>(row, "Kontorkostnader");	// 6800 

                    accountingItem.CostOfData = ExcelUtils.GetField<decimal>(row, "Datakostnader");	// 6810 
                    accountingItem.CostOfPhoneInternetUse = ExcelUtils.GetField<decimal>(row, "Telefon Internett Bruk");	// 6900
                    accountingItem.PrivateUseOfECom = ExcelUtils.GetField<decimal>(row, "Privat bruk av el.kommunikasjon");    // 7098

                    accountingItem.CostForTravelAndAllowance = ExcelUtils.GetField<decimal>(row, "Reise og Diett");	// 7140
                    accountingItem.CostOfAdvertising = ExcelUtils.GetField<decimal>(row, "Reklamekostnader");	// 7330
                    accountingItem.CostOfOther = ExcelUtils.GetField<decimal>(row, "Diverse annet");	// 7700

                    accountingItem.FeesBank = ExcelUtils.GetField<decimal>(row, "Gebyrer Bank");	// 7770
                    accountingItem.FeesPaypal = ExcelUtils.GetField<decimal>(row, "Gebyrer Paypal");	// 7780
                    accountingItem.FeesStripe = ExcelUtils.GetField<decimal>(row, "Gebyrer Stripe");	// 7785 

                    accountingItem.CostForEstablishment = ExcelUtils.GetField<decimal>(row, "Etableringskostnader");	// 7790

                    accountingItem.IncomeFinance = ExcelUtils.GetField<decimal>(row, "Finansinntekter");	// 8099
                    accountingItem.CostOfFinance = ExcelUtils.GetField<decimal>(row, "Finanskostnader");	// 8199

                    accountingItem.Investments = ExcelUtils.GetField<decimal>(row, "Investeringer");	// 1200
                    accountingItem.AccountsReceivable = ExcelUtils.GetField<decimal>(row, "Kundefordringer");	// 1500
                    accountingItem.PersonalWithdrawal = ExcelUtils.GetField<decimal>(row, "Privat uttak");
                    accountingItem.PersonalDeposit = ExcelUtils.GetField<decimal>(row, "Privat innskudd");
                    // cSpell:enable
                    existingAccountingItems.Add(row, accountingItem);
                }

                // reduce the old Accounting Spreadsheet and remove the entries that doesn't have a number
                var existingAccountingItemsToUpdate =
                    (from row in existingAccountingItems
                     where
                     row.Value.Gateway == "stripe"
                     || row.Value.Gateway == "paypal"
                     orderby row.Value.Number ascending
                     select row);

                int updateRowCounter = 0;
                int updateRowTotalCount = existingAccountingItemsToUpdate.Count();
                Console.Out.WriteLine("Updating {0} rows", updateRowTotalCount);
                foreach (var updateRow in existingAccountingItemsToUpdate)
                {
                    updateRowCounter++;
                    Console.Out.Write("\rUpdating row {0}/{1} ({2})", updateRowCounter, updateRowTotalCount, updateRow.Key.RangeAddress);

                    // find match
                    var result = (from a in newAccountingItems
                                  where a.ArchiveReference == updateRow.Value.ArchiveReference
                                  && a.Date == updateRow.Value.Date
                                  && a.Text == updateRow.Value.Text
                                  && a.AccountPaypal == updateRow.Value.AccountPaypal
                                  && a.AccountStripe == updateRow.Value.AccountStripe
                                  select a).ToList();
                    if (result.Count() > 0)
                    {
                        if (result.Count() > 1)
                        {
                            // error
                            Console.Out.Write("\rFailed finding only single matching accounting entries to update from {0} ...", updateRow.Key.RangeAddress);
                            return;
                        }
                        else
                        {
                            var matchingAccountingElement = result.First();
                            updateRow.Key.Cell(6).Value = matchingAccountingElement.TransactionID;
                            updateRow.Key.Cell(11).Value = matchingAccountingElement.ErrorMessage;
                        }
                    }
                    else
                    {
                        // error, none found
                        Console.Out.Write("\rFailed finding matching accounting entry to update from {0} ...", updateRow.Key.RangeAddress);
                    }
                }
            }

            // resize
            //ws.Columns().AdjustToContents();  // Adjust column width
            //ws.Rows().AdjustToContents();     // Adjust row heights

            wb.Save();
            Console.Out.WriteLine("\rSuccessfully updated accounting file!");
        }

        static void SetExcelRowFormulas(IXLRangeRow row)
        {
            int currentRow = row.RowNumber();

            // create formulas
            string controlFormula = string.Format("=IF(BA{0}=0,\" \",\"!!FEIL!!\")", currentRow);
            string sumPreRoundingFormula = string.Format("=SUM(Q{0}:AY{0})", currentRow);
            string sumRoundedFormula = string.Format("=ROUND(AZ{0},2)", currentRow);
            string vatSales = string.Format("=-(O{0}/1.25)*0.25", currentRow);
            string salesVATExempt = string.Format("=-(O{0}/1.25)", currentRow);

            // apply formulas to cells.
            row.Cell("A").FormulaA1 = controlFormula;
            row.Cell("AZ").FormulaA1 = sumPreRoundingFormula;
            row.Cell("BA").FormulaA1 = sumRoundedFormula;

            // add VAT formulas
            if (row.Cell("P").Value.Equals("NOK")
                && (row.Cell("H").Value.Equals("SHOPIFY"))
                && (row.Cell("X").GetValue<decimal>() != 0))
            {
                row.Cell("V").FormulaA1 = vatSales;
                row.Cell("X").FormulaA1 = salesVATExempt;
            }
        }

        static void SetExcelRowStyles(IXLRangeRow row)
        {
            int currentRow = row.RowNumber();

            // set font color for control column
            row.Cell("A").Style.Font.FontColor = XLColor.Red;
            row.Cell("A").Style.Font.Bold = true;

            // set background color for VAT
            var lightGreen = XLColor.FromArgb(0xD8E4BC);
            var lighterGreen = XLColor.FromArgb(0xEBF1DE);
            var green = currentRow % 2 == 0 ? lightGreen : lighterGreen;
            row.Cells("U", "V").Style.Fill.BackgroundColor = green;

            // set background color for investments, withdrawal and deposits
            var lightBlue = XLColor.FromArgb(0xEAF1FA);
            var lighterBlue = XLColor.FromArgb(0xC5D9F1);
            var blue = currentRow % 2 == 0 ? lightBlue : lighterBlue;
            row.Cells("AV", "AY").Style.Fill.BackgroundColor = blue;

            // set background color for control sum
            var lightRed = XLColor.FromArgb(0xE6B8B7);
            var lighterRed = XLColor.FromArgb(0xF2DCDB);
            var red = currentRow % 2 == 0 ? lightRed : lighterRed;
            row.Cell("BA").Style.Fill.BackgroundColor = red;

            // set column formats
            row.Cell("C").Style.NumberFormat.Format = "dd.MM.yyyy";

            // Arkivreferanse or ArchiveReference has so many digits 
            // that Excel will truncate it, therefore we need to ensure
            // that the long is stored as text and not a number
            row.Cell("E").DataType = XLDataType.Text;

            // Custom formats for numbers in Excel are entered in this format:
            // positive number format;negative number format;zero format;text format
            row.Cell("O").Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00;";
            row.Cell("O").DataType = XLDataType.Number;

            // set style and format for the decimal range
            row.Cells("Q", "BA").Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00;";
            row.Cells("Q", "BA").DataType = XLDataType.Number;
        }

        static void SetExcelTableTotalsRowFunction(IXLTable table)
        {
            table.ShowTotalsRow = true;

            // set sum functions for each of the table columns 
            // cSpell:disable
            table.Field("Paypal").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 1910
            table.Field("Stripe").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 1915
            table.Field("Vipps").TotalsRowFunction = XLTotalsRowFunction.Sum;               // 1918
            table.Field("Bank").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 1920

            table.Field("MVA Kjøp").TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.Field("MVA Salg").TotalsRowFunction = XLTotalsRowFunction.Sum;

            table.Field("Oppgjørskonto mva").TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.Field("Salg mva-pliktig").TotalsRowFunction = XLTotalsRowFunction.Sum;                   // 3000
            table.Field("Salg avgiftsfritt").TotalsRowFunction = XLTotalsRowFunction.Sum;             // 3100

            table.Field("Varekostnad").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 4005
            table.Field("Forbruk for videresalg").TotalsRowFunction = XLTotalsRowFunction.Sum;           // 4300
            table.Field("Lønn").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 5000
            table.Field("Arb.giver avgift").TotalsRowFunction = XLTotalsRowFunction.Sum;           // 5400
            table.Field("Avskrivninger").TotalsRowFunction = XLTotalsRowFunction.Sum;        // 6000
            table.Field("Frakt").TotalsRowFunction = XLTotalsRowFunction.Sum;            // 6100
            table.Field("Strøm").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 6340 
            table.Field("Verktøy inventar").TotalsRowFunction = XLTotalsRowFunction.Sum;      // 6500
            table.Field("Vedlikehold").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 6695
            table.Field("Kontorkostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;          // 6800 

            table.Field("Datakostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 6810 
            table.Field("Telefon Internett Bruk").TotalsRowFunction = XLTotalsRowFunction.Sum;        // 6900
            table.Field("Privat bruk av el.kommunikasjon").TotalsRowFunction = XLTotalsRowFunction.Sum;        // 7098
            table.Field("Reise og Diett").TotalsRowFunction = XLTotalsRowFunction.Sum;  // 7140
            table.Field("Reklamekostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;          // 7330
            table.Field("Diverse annet").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 7700

            table.Field("Gebyrer Bank").TotalsRowFunction = XLTotalsRowFunction.Sum;                   // 7770
            table.Field("Gebyrer Paypal").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 7780
            table.Field("Gebyrer Stripe").TotalsRowFunction = XLTotalsRowFunction.Sum;                 // 7785 

            table.Field("Etableringskostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;       // 7790

            table.Field("Finansinntekter").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 8099
            table.Field("Finanskostnader").TotalsRowFunction = XLTotalsRowFunction.Sum;              // 8199

            table.Field("Investeringer").TotalsRowFunction = XLTotalsRowFunction.Sum;                // 1200
            table.Field("Kundefordringer").TotalsRowFunction = XLTotalsRowFunction.Sum;         // 1500
            table.Field("Privat uttak").TotalsRowFunction = XLTotalsRowFunction.Sum;
            table.Field("Privat innskudd").TotalsRowFunction = XLTotalsRowFunction.Sum;
            // cSpell:enable
        }

        static DataTable GetDataTable(List<AccountingItem> accountingItems)
        {
            DataTable dt = new DataTable();
            // cSpell:disable            
            dt.Columns.Add("Kontroll", typeof(string));

            dt.Columns.Add("Periode", typeof(int));
            dt.Columns.Add("Dato", typeof(DateTime));
            dt.Columns.Add("Bilagsnr.", typeof(int));
            dt.Columns.Add("Arkivreferanse", typeof(string)); // ensure the archive reference is stored as text
            dt.Columns.Add("TransaksjonsId", typeof(string));
            dt.Columns.Add("Type", typeof(string));
            dt.Columns.Add("Regnskapstype", typeof(string));
            dt.Columns.Add("Tekst", typeof(string));
            dt.Columns.Add("Kundenavn", typeof(string));
            dt.Columns.Add("Feilmelding", typeof(string));
            dt.Columns.Add("Gateway", typeof(string));
            dt.Columns.Add("Num Salg", typeof(string));
            dt.Columns.Add("Num Kjøp", typeof(string));
            dt.Columns.Add("Kjøp annen valuta", typeof(decimal));
            dt.Columns.Add("Annen valuta", typeof(string));

            dt.Columns.Add("Paypal", typeof(decimal));                          // 1910
            dt.Columns.Add("Stripe", typeof(decimal));                          // 1912
            dt.Columns.Add("Vipps", typeof(decimal));                           // 1914
            dt.Columns.Add("Bank", typeof(decimal));                            // 1920

            dt.Columns.Add("MVA Kjøp", typeof(decimal));
            dt.Columns.Add("MVA Salg", typeof(decimal));

            dt.Columns.Add("Oppgjørskonto mva", typeof(decimal));               // 2740
            dt.Columns.Add("Salg mva-pliktig", typeof(decimal));                // 3000
            dt.Columns.Add("Salg avgiftsfritt", typeof(decimal));               // 3100

            dt.Columns.Add("Varekostnad", typeof(decimal));                     // 4005
            dt.Columns.Add("Forbruk for videresalg", typeof(decimal));          // 4300
            dt.Columns.Add("Lønn", typeof(decimal));                            // 5000
            dt.Columns.Add("Arb.giver avgift", typeof(decimal));                // 5400
            dt.Columns.Add("Avskrivninger", typeof(decimal));                   // 6000
            dt.Columns.Add("Frakt", typeof(decimal));                           // 6100
            dt.Columns.Add("Strøm", typeof(decimal));                           // 6340 
            dt.Columns.Add("Verktøy inventar", typeof(decimal));                // 6500
            dt.Columns.Add("Vedlikehold", typeof(decimal));                     // 6695
            dt.Columns.Add("Kontorkostnader", typeof(decimal));                 // 6800 

            dt.Columns.Add("Datakostnader", typeof(decimal));                   // 6810 
            dt.Columns.Add("Telefon Internett Bruk", typeof(decimal));          // 6900
            dt.Columns.Add("Privat bruk av el.kommunikasjon", typeof(decimal)); // 7098
            dt.Columns.Add("Reise og Diett", typeof(decimal));                  // 7140
            dt.Columns.Add("Reklamekostnader", typeof(decimal));                // 7330
            dt.Columns.Add("Diverse annet", typeof(decimal));                   // 7700

            dt.Columns.Add("Gebyrer Bank", typeof(decimal));                    // 7770
            dt.Columns.Add("Gebyrer Paypal", typeof(decimal));                  // 7780
            dt.Columns.Add("Gebyrer Stripe", typeof(decimal));                  // 7785 

            dt.Columns.Add("Etableringskostnader", typeof(decimal));            // 7790

            dt.Columns.Add("Finansinntekter", typeof(decimal));                 // 8099
            dt.Columns.Add("Finanskostnader", typeof(decimal));                 // 8199

            dt.Columns.Add("Investeringer", typeof(decimal));                   // 1200
            dt.Columns.Add("Kundefordringer", typeof(decimal));                 // 1500
            dt.Columns.Add("Privat uttak", typeof(decimal));
            dt.Columns.Add("Privat innskudd", typeof(decimal));

            dt.Columns.Add("Sum før avrunding", typeof(decimal));
            dt.Columns.Add("Sum", typeof(decimal));
            // cSpell:enable

            foreach (var accountingItem in accountingItems)
            {
                dt.Rows.Add(
                    "",
                    accountingItem.Periode,
                    accountingItem.Date,
                    accountingItem.Number,
                    accountingItem.ArchiveReference,
                    accountingItem.TransactionID,
                    accountingItem.Type,
                    accountingItem.AccountingType,
                    accountingItem.Text,
                    accountingItem.CustomerName,
                    accountingItem.ErrorMessage,
                    accountingItem.Gateway,
                    accountingItem.NumSale,
                    accountingItem.NumPurchase,
                    accountingItem.PurchaseOtherCurrency,
                    accountingItem.OtherCurrency,

                    accountingItem.AccountPaypal,               // 1910
                    accountingItem.AccountStripe,               // 1915
                    accountingItem.AccountVipps,                // 1918
                    accountingItem.AccountBank,                 // 1920

                    accountingItem.VATPurchase,
                    accountingItem.VATSales,

                    accountingItem.VATSettlementAccount,        // 2740
                    accountingItem.SalesVAT,                    // 3000
                    accountingItem.SalesVATExempt,              // 3100

                    accountingItem.CostOfGoods,                 // 4005
                    accountingItem.CostForReselling,            // 4300
                    accountingItem.CostForSalary,               // 5000
                    accountingItem.CostForSalaryTax,            // 5400
                    accountingItem.CostForDepreciation,         // 6000
                    accountingItem.CostForShipping,             // 6100
                    accountingItem.CostForElectricity,          // 6340 
                    accountingItem.CostForToolsInventory,       // 6500
                    accountingItem.CostForMaintenance,          // 6695
                    accountingItem.CostForFacilities,           // 6800 

                    accountingItem.CostOfData,                  // 6810 
                    accountingItem.CostOfPhoneInternetUse,      // 6900
                    accountingItem.PrivateUseOfECom,            // 7098
                    accountingItem.CostForTravelAndAllowance,   // 7140
                    accountingItem.CostOfAdvertising,           // 7330
                    accountingItem.CostOfOther,                 // 7700

                    accountingItem.FeesBank,                    // 7770
                    accountingItem.FeesPaypal,                  // 7780
                    accountingItem.FeesStripe,                  // 7785 

                    accountingItem.CostForEstablishment,        // 7790

                    accountingItem.IncomeFinance,               // 8099
                    accountingItem.CostOfFinance,               // 8199

                    accountingItem.Investments,                 // 1200
                    accountingItem.AccountsReceivable,          // 1500
                    accountingItem.PersonalWithdrawal,
                    accountingItem.PersonalDeposit
                    );
            }

            return dt;
        }
        #endregion

        static List<AccountingItem> ProcessBankAccountStatement(IMyConfiguration configuration, SkandiabankenBankStatement skandiabankenBankStatement, List<string> customerNames, List<StripeTransaction> stripeTransactions, List<PayPalTransaction> paypalTransactions)
        {
            bool PROCESS_ALIEXPRESS = false;

            var accountingList = new List<AccountingItem>();

            if (skandiabankenBankStatement == null) return accountingList;

            var date = new Date();
            var from = date.FirstDayOfTheYear;
            var to = date.CurrentDate;

            // prepopulate some lookup lists
            var stripePayoutTransactions = StripePayoutFactory.Instance.GetLatest(configuration, false);
            Console.Out.WriteLine("Successfully read Stripe payout transactions ...");

            // check if we are processing aliexpress
            var oberloOrders = new List<OberloOrder>();
            var aliExpressOrderGroups = new List<AliExpressOrderGroup>();
            if (PROCESS_ALIEXPRESS)
            {
                // process oberlo
                oberloOrders = OberloFactory.Instance.GetLatest(configuration);

                // and then ali express
                var aliExpressOrders = AliExpressFactory.Instance.GetLatest(configuration);
                aliExpressOrderGroups = AliExpress.CombineOrders(aliExpressOrders);
            }

            // run through the bank account transactions
            var skandiabankenTransactions = skandiabankenBankStatement.Transactions;

            // define hashes so we can track used order numbers and transaction Ids
            var usedStripePayoutTransactionIDs = new HashSet<string>();
            var usedOrderNumbers = new HashSet<string>();

            // and map each one to the right meta information
            foreach (var skandiabankenTransaction in skandiabankenTransactions)
            {
                // define accounting item
                var accountingItem = new AccountingItem();

                // set date to closer to midnight (sorts better)
                accountingItem.Date = new DateTime(
                    skandiabankenTransaction.TransactionDate.Year,
                    skandiabankenTransaction.TransactionDate.Month,
                    skandiabankenTransaction.TransactionDate.Day,
                    23, 59, 00);

                accountingItem.ArchiveReference = skandiabankenTransaction.ArchiveReference;

                if (accountingItem.ArchiveReference.Equals("fddfd41eb41537643cf826b34398d632"))
                {
                    // breakpoint here
                }

                // extract properties from the transaction text
                skandiabankenTransaction.ExtractAccountingInformationAPI();
                var accountingType = skandiabankenTransaction.AccountingType;
                accountingItem.AccountingType = skandiabankenTransaction.GetAccountingTypeString();
                accountingItem.Type = skandiabankenTransaction.Type;

                // 1. If purchase or return from purchase 
                if (skandiabankenTransaction.Type.Equals("VISA VARE") && (
                    accountingType == SBankenTransaction.AccountingTypeEnum.CostOfWebShop ||
                    accountingType == SBankenTransaction.AccountingTypeEnum.CostOfAdvertising ||
                    accountingType == SBankenTransaction.AccountingTypeEnum.CostOfDomain ||
                    accountingType == SBankenTransaction.AccountingTypeEnum.CostOfServer ||
                    accountingType == SBankenTransaction.AccountingTypeEnum.IncomeReturn))
                {

                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0:dd.MM.yyyy} {1} {2} {3:C} (Kurs: {4})", skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor, skandiabankenTransaction.ExternalPurchaseAmount, skandiabankenTransaction.ExternalPurchaseCurrency, skandiabankenTransaction.ExternalPurchaseExchangeRate);
                    accountingItem.PurchaseOtherCurrency = skandiabankenTransaction.ExternalPurchaseAmount;
                    accountingItem.OtherCurrency = skandiabankenTransaction.ExternalPurchaseCurrency.ToUpper();
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;

                    switch (accountingType)
                    {
                        case SBankenTransaction.AccountingTypeEnum.CostOfWebShop:
                        case SBankenTransaction.AccountingTypeEnum.CostOfDomain:
                        case SBankenTransaction.AccountingTypeEnum.CostOfServer:
                            accountingItem.CostOfData = -skandiabankenTransaction.AccountChange;
                            break;
                        case SBankenTransaction.AccountingTypeEnum.CostOfAdvertising:
                            accountingItem.CostOfAdvertising = -skandiabankenTransaction.AccountChange;
                            break;
                    }
                }

                // 1. If AliExpress or PayPal purchase
                else if (skandiabankenTransaction.Type.Equals("VISA VARE") &&
                    accountingType == SBankenTransaction.AccountingTypeEnum.CostOfGoods)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0:dd.MM.yyyy} {1} {2} {3:C} (Kurs: {4})", skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor, skandiabankenTransaction.ExternalPurchaseAmount, skandiabankenTransaction.ExternalPurchaseCurrency, skandiabankenTransaction.ExternalPurchaseExchangeRate);
                    accountingItem.PurchaseOtherCurrency = skandiabankenTransaction.ExternalPurchaseAmount;
                    accountingItem.OtherCurrency = skandiabankenTransaction.ExternalPurchaseCurrency.ToUpper();
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                    accountingItem.CostForReselling = -skandiabankenTransaction.AccountChange;

                    if (PROCESS_ALIEXPRESS && skandiabankenTransaction.ExternalPurchaseVendor.CaseInsensitiveContains("AliExpress"))
                    {
                        FindAliExpressOrderNumber(usedOrderNumbers, aliExpressOrderGroups, oberloOrders, skandiabankenTransaction, accountingItem);
                    }
                }

                // 2. Transfer Paypal
                else if (accountingType == SBankenTransaction.AccountingTypeEnum.TransferPaypal)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0:dd.MM.yyyy} {1}", skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor);
                    accountingItem.Gateway = "paypal";

                    accountingItem.AccountPaypal = -skandiabankenTransaction.AccountChange;
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;

                    // lookup the paypal transaction
                    var startDate = skandiabankenTransaction.ExternalPurchaseDate.AddDays(-3);
                    var endDate = skandiabankenTransaction.ExternalPurchaseDate.AddDays(1);

                    var paypalQuery =
                    from transaction in paypalTransactions
                    let grossAmount = transaction.GrossAmount
                    let timestamp = transaction.Timestamp
                    where
                    transaction.Type.Equals("Transfer")
                    && (grossAmount == -skandiabankenTransaction.AccountChange)
                    && (timestamp.Date >= startDate.Date && timestamp.Date <= endDate.Date)
                    orderby timestamp ascending
                    select transaction;

                    if (paypalQuery.Count() > 1)
                    {
                        // more than one transaction found ?!
                        Console.Out.WriteLine("ERROR: FOUND MORE THAN ONE PAYPAL PAYOUT!");
                        accountingItem.ErrorMessage = "Paypal: More than one payout found, choose one";
                    }
                    else if (paypalQuery.Count() > 0)
                    {
                        // one match
                        var paypalTransaction = paypalQuery.First();

                        // store the transaction id
                        accountingItem.TransactionID = paypalTransaction.TransactionID;
                    }
                    else
                    {
                        Console.Out.WriteLine("ERROR: NO PAYPAL PAYOUTS FOR {0:C} FOUND BETWEEN {1:dd.MM.yyyy} and {2:dd.MM.yyyy}!", skandiabankenTransaction.AccountChange, startDate, endDate);
                        accountingItem.ErrorMessage = "Paypal: No payouts found";
                    }
                }

                // 3. Transfer Stripe
                else if (accountingType == SBankenTransaction.AccountingTypeEnum.TransferStripe)
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0:dd.MM.yyyy} {1}", skandiabankenTransaction.ExternalPurchaseDate, skandiabankenTransaction.ExternalPurchaseVendor);
                    accountingItem.Gateway = "stripe";

                    accountingItem.AccountStripe = -skandiabankenTransaction.AccountChange;
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;

                    // lookup the stripe payout transaction
                    var startDate = skandiabankenTransaction.ExternalPurchaseDate.AddDays(-4);
                    var endDate = skandiabankenTransaction.ExternalPurchaseDate.AddDays(4);

                    var stripeQuery =
                    from transaction in stripePayoutTransactions
                    where
                    transaction.Paid &&
                    transaction.Amount == skandiabankenTransaction.AccountChange &&
                     (transaction.AvailableOn.Date >= startDate.Date && transaction.AvailableOn.Date <= endDate.Date)
                    orderby transaction.Created ascending
                    select transaction;

                    if (stripeQuery.Count() > 1)
                    {
                        Console.WriteLine("\tSEVERAL MATCHING STRIPE PAYOUTS FOUND ...");

                        bool notFound = true;
                        foreach (var item in stripeQuery.Reverse())
                        {
                            var stripePayoutTransactionID = item.TransactionID;
                            if (!usedStripePayoutTransactionIDs.Contains(stripePayoutTransactionID))
                            {
                                notFound = false;
                                usedStripePayoutTransactionIDs.Add(stripePayoutTransactionID);
                                accountingItem.TransactionID = stripePayoutTransactionID;
                                Console.WriteLine("\tSELECTED: {0} {1}", accountingItem.TransactionID, accountingItem.Text);
                                break;
                            }
                        }

                        if (notFound)
                        {
                            Console.Out.WriteLine("ERROR: COULD NOT FIND MATCHING STRIPE PAYOUT!");
                            accountingItem.ErrorMessage = "Stripe: Could not find matching payout";
                        }
                    }
                    else if (stripeQuery.Count() > 0)
                    {
                        // one match
                        var stripeTransaction = stripeQuery.First();

                        // store the transaction id
                        accountingItem.TransactionID = stripeTransaction.TransactionID;
                    }
                    else
                    {
                        Console.Out.WriteLine("ERROR: NO STRIPE PAYOUT FOR {0:C} FOUND BETWEEN {1:dd.MM.yyyy} and {2:dd.MM.yyyy}!", skandiabankenTransaction.AccountChange, startDate, endDate);
                        accountingItem.ErrorMessage = "Stripe: No payouts found";
                    }
                }

                else if (customerNames.Contains(skandiabankenTransaction.Text))
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);
                    accountingItem.Text = string.Format("{0}", skandiabankenTransaction.Text);
                    accountingItem.Gateway = "vipps";
                    accountingItem.AccountingType = "OVERFØRSEL VIPPS";
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;
                    accountingItem.AccountVipps = -skandiabankenTransaction.AccountChange;
                }

                // 4. None of those above
                else
                {
                    Console.WriteLine("{0}", skandiabankenTransaction);

                    // if the text contains an USD pattern, use it
                    Regex usdPattern = new Regex(@"USD\s+(\d+(\.\d+)?)", RegexOptions.Compiled);
                    var matchUSD = usdPattern.Match(skandiabankenTransaction.Text);
                    if (matchUSD.Success)
                    {
                        var purchaseOtherCurrencyString = matchUSD.Groups[1].Value.ToString();
                        decimal purchaseOtherCurrency;
                        decimal.TryParse(purchaseOtherCurrencyString, NumberStyles.Currency, CultureInfo.InvariantCulture, out purchaseOtherCurrency);
                        accountingItem.PurchaseOtherCurrency = purchaseOtherCurrency;
                        accountingItem.OtherCurrency = "USD";
                    }

                    accountingItem.Text = string.Format("{0}", skandiabankenTransaction.Text);
                    accountingItem.AccountBank = skandiabankenTransaction.AccountChange;

                    switch (accountingType)
                    {
                        case SBankenTransaction.AccountingTypeEnum.CostOfWebShop:
                        case SBankenTransaction.AccountingTypeEnum.CostOfDomain:
                        case SBankenTransaction.AccountingTypeEnum.CostOfServer:
                            accountingItem.CostOfData = -skandiabankenTransaction.AccountChange;
                            break;
                        case SBankenTransaction.AccountingTypeEnum.CostOfAdvertising:
                            accountingItem.CostOfAdvertising = -skandiabankenTransaction.AccountChange;
                            break;
                        case SBankenTransaction.AccountingTypeEnum.CostOfTryouts:
                            accountingItem.CostOfGoods = -skandiabankenTransaction.AccountChange;
                            break;
                        case SBankenTransaction.AccountingTypeEnum.CostOfBank:
                            accountingItem.CostOfFinance = -skandiabankenTransaction.AccountChange;
                            break;
                        case SBankenTransaction.AccountingTypeEnum.IncomeInterest:
                            accountingItem.IncomeFinance = -skandiabankenTransaction.AccountChange;
                            break;
                        case SBankenTransaction.AccountingTypeEnum.IncomeReturn:
                            accountingItem.CostForReselling = -skandiabankenTransaction.AccountChange;
                            break;
                        case SBankenTransaction.AccountingTypeEnum.IncomeVATReturn:
                            accountingItem.VATSettlementAccount = -skandiabankenTransaction.AccountChange;
                            break;
                        case SBankenTransaction.AccountingTypeEnum.CostOfVAT:
                            accountingItem.VATSettlementAccount = -skandiabankenTransaction.AccountChange;
                            accountingItem.ErrorMessage = "Please add VAT payment period";
                            break;
                    }
                }

                accountingList.Add(accountingItem);
            }
            return accountingList;
        }

        static List<AccountingItem> ProcessShopifyStatement(IMyConfiguration configuration, List<string> customerNames, List<StripeTransaction> stripeTransactions, List<PayPalTransaction> paypalTransactions)
        {
            var accountingList = new List<AccountingItem>();

            // get shopify configuration parameters
            string shopifyDomain = configuration.GetValue("ShopifyDomain");
            string shopifyAPIKey = configuration.GetValue("ShopifyAPIKey");
            string shopifyAPIPassword = configuration.GetValue("ShopifyAPIPassword");

            // add date filter, created_at_min and created_at_max
            var date = new Date();
            var from = date.FirstDayOfTheYear; //.AddDays(-30); // always go back a month
            var to = date.CurrentDate;
            string querySuffix = string.Format(CultureInfo.InvariantCulture, "status=any&created_at_min={0:yyyy-MM-ddTHH:mm:sszzz}&created_at_max={1:yyyy-MM-ddTHH:mm:sszzz}", from, to);
            var shopifyOrders = Shopify.ReadShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword, querySuffix);
            Console.Out.WriteLine("Successfully read all Shopify orders ...");

            Console.Out.WriteLine("Processing Shopify orders started ...");
            foreach (var shopifyOrder in shopifyOrders)
            {
                // skip, not paid (pending), cancelled (voided) and fully refunded orders (refunded)
                if (shopifyOrder.FinancialStatus.Equals("refunded_IGNORE")
                    || shopifyOrder.FinancialStatus.Equals("voided")
                    || shopifyOrder.FinancialStatus.Equals("pending")
                    || shopifyOrder.Cancelled == true
                    ) continue;

                if (shopifyOrder.FinancialStatus.Equals("refunded"))
                {
                    //
                }

                // define accounting item
                var accountingItem = new AccountingItem();
                accountingItem.Date = shopifyOrder.CreatedAt;
                accountingItem.ArchiveReference = shopifyOrder.Id.ToString();
                accountingItem.Type = string.Format("{0} {1}", shopifyOrder.FinancialStatus, shopifyOrder.FulfillmentStatus);
                accountingItem.AccountingType = "SHOPIFY";
                accountingItem.Text = string.Format("SALG {0} {1}", shopifyOrder.CustomerName, shopifyOrder.PaymentId);
                accountingItem.CustomerName = shopifyOrder.CustomerName;

                // add to customer name list
                customerNames.Add(accountingItem.CustomerName);

                if (shopifyOrder.Gateway != null)
                {
                    accountingItem.Gateway = shopifyOrder.Gateway.ToLower();
                }
                accountingItem.NumSale = shopifyOrder.Name;

                var startDate = shopifyOrder.ProcessedAt.AddDays(-1);
                var endDate = shopifyOrder.ProcessedAt.AddDays(1);

                switch (accountingItem.Gateway)
                {
                    case "vipps":
                        accountingItem.PurchaseOtherCurrency = shopifyOrder.TotalPrice;
                        accountingItem.OtherCurrency = "NOK";

                        //accountingItem.FeesVipps = fee;
                        accountingItem.AccountVipps = shopifyOrder.TotalPrice;

                        break;
                    case "stripe":

                        accountingItem.PurchaseOtherCurrency = shopifyOrder.TotalPrice;
                        accountingItem.OtherCurrency = "NOK";

                        // lookup the stripe transaction
                        var stripeQuery =
                        from transaction in stripeTransactions
                        where
                        transaction.Paid &&
                        transaction.CustomerEmail.Equals(shopifyOrder.CustomerEmail) &&
                        transaction.Amount == shopifyOrder.TotalPrice &&
                         (transaction.Created.Date >= startDate.Date && transaction.Created.Date <= endDate.Date)
                        orderby transaction.Created ascending
                        select transaction;

                        if (stripeQuery.Count() > 1)
                        {
                            // more than one ?!
                            Console.Out.WriteLine("ERROR: FOUND MORE THAN ONE MATCHING STRIPE TRANSACTION!");
                            accountingItem.ErrorMessage = "Stripe: More than one found, choose one";
                        }
                        else if (stripeQuery.Count() > 0)
                        {
                            // one match
                            var stripeTransaction = stripeQuery.First();
                            decimal amount = stripeTransaction.Amount;
                            decimal net = stripeTransaction.Net;
                            decimal fee = stripeTransaction.Fee;

                            accountingItem.FeesStripe = fee;
                            accountingItem.AccountStripe = net;

                            // also store the transaction id
                            accountingItem.TransactionID = stripeTransaction.TransactionID;
                        }
                        else
                        {
                            Console.Out.WriteLine("ERROR: NO STRIPE TRANSACTIONS FOR {0:C} FOUND FOR {1} {2} BETWEEN {3:dd.MM.yyyy} and {4:dd.MM.yyyy}!", shopifyOrder.TotalPrice, shopifyOrder.Name, shopifyOrder.CustomerName, startDate, endDate);
                            accountingItem.ErrorMessage = "Stripe: No transactions found";
                        }

                        break;
                    case "paypal":

                        accountingItem.PurchaseOtherCurrency = shopifyOrder.TotalPrice;
                        accountingItem.OtherCurrency = "NOK";

                        // lookup the paypal transaction
                        var paypalQuery =
                        from transaction in paypalTransactions
                        let grossAmount = transaction.GrossAmount
                        let timestamp = transaction.Timestamp
                        where
                        transaction.Status.Equals("Completed")
                        //&& (null != transaction.Payer && transaction.Payer.Equals(shopifyOrder.CustomerEmail))
                        && (
                        (null != transaction.PayerDisplayName && transaction.PayerDisplayName.Equals(shopifyOrder.CustomerName, StringComparison.InvariantCultureIgnoreCase))
                        ||
                        (null != transaction.Payer && transaction.Payer.Equals(shopifyOrder.CustomerEmail, StringComparison.InvariantCultureIgnoreCase))
                        )
                        && (grossAmount == shopifyOrder.TotalPrice)
                        && (timestamp.Date >= startDate.Date && timestamp.Date <= endDate.Date)
                        //&& (timestamp.Date == shopifyOrder.ProcessedAt.Date)
                        orderby timestamp ascending
                        select transaction;

                        if (paypalQuery.Count() > 1)
                        {
                            // more than one ?!
                            Console.Out.WriteLine("ERROR: FOUND MORE THAN ONE PAYPAL TRANSACTION!");
                            accountingItem.ErrorMessage = "Paypal: More than one found, choose one";
                        }
                        else if (paypalQuery.Count() > 0)
                        {
                            // one match
                            var paypalTransaction = paypalQuery.First();
                            decimal amount = paypalTransaction.GrossAmount;
                            decimal net = paypalTransaction.NetAmount;
                            decimal fee = paypalTransaction.FeeAmount;

                            accountingItem.FeesPaypal = -fee;
                            accountingItem.AccountPaypal = net;

                            // also store the transaction id
                            accountingItem.TransactionID = paypalTransaction.TransactionID;
                        }
                        else
                        {
                            Console.Out.WriteLine("ERROR: NO PAYPAL TRANSACTIONS FOR {0:C} FOUND FOR {1} {2} BETWEEN {3:dd.MM.yyyy} and {4:dd.MM.yyyy}!", shopifyOrder.TotalPrice, shopifyOrder.Name, shopifyOrder.CustomerName, startDate, endDate);
                            accountingItem.ErrorMessage = "Paypal: No transactions found";
                        }

                        break;
                }

                // fix VAT
                if (shopifyOrder.TotalTax != 0)
                {
                    accountingItem.SalesVAT = -(shopifyOrder.TotalPrice / (decimal)1.25);
                    accountingItem.VATSales = accountingItem.SalesVAT * (decimal)0.25;
                }
                else
                {
                    accountingItem.SalesVATExempt = -shopifyOrder.TotalPrice;
                }

                // check if free gift
                if (shopifyOrder.TotalPrice == 0)
                {
                    accountingItem.AccountingType += " FREE";
                    accountingItem.Gateway = "none";
                }

                accountingList.Add(accountingItem);
            }

            return accountingList;
        }

        #region AliExpress Methods
        static void FindAliExpressOrderNumber(HashSet<string> usedOrderNumbers, List<AliExpressOrderGroup> aliExpressOrderGroups, List<OberloOrder> oberloOrders, SBankenTransaction skandiabankenTransaction, AccountingItem accountingItem)
        {
            // set start and stop date
            var startDate = skandiabankenTransaction.ExternalPurchaseDate.AddDays(-4);
            var endDate = skandiabankenTransaction.ExternalPurchaseDate.AddDays(2);

            // lookup in AliExpress purchase list
            // matching ordertime and orderamount
            var aliExpressQuery =
                from order in aliExpressOrderGroups
                where
                (order.OrderTime.Date >= startDate.Date && order.OrderTime.Date <= endDate.Date) &&
                order.OrderAmount == skandiabankenTransaction.ExternalPurchaseAmount
                orderby order.OrderTime ascending
                select order;

            // if the count is more than one, we cannot match easily 
            if (aliExpressQuery.Count() > 1)
            {
                // first check if one of the found orders was ordered on the given purchase date
                var aliExpressQueryExactDate =
                from order in aliExpressQuery
                where
                order.OrderTime.Date == skandiabankenTransaction.ExternalPurchaseDate.Date
                orderby order.OrderTime ascending
                select order;

                // if the count is only one, we have a single match
                if (aliExpressQueryExactDate.Count() == 1)
                {
                    ProcessAliExpressMatch(usedOrderNumbers, aliExpressQueryExactDate, oberloOrders, accountingItem);
                    return;
                }
                // use the original query and present the results
                else
                {
                    ProcessAliExpressMatch(usedOrderNumbers, aliExpressQuery, oberloOrders, accountingItem);
                }
            }
            // if the count is only one, we have a single match
            else if (aliExpressQuery.Count() == 1)
            {
                ProcessAliExpressMatch(usedOrderNumbers, aliExpressQuery, oberloOrders, accountingItem);
            }
            // no orders found
            else
            {
                // could not find shopify order numbers
                Console.WriteLine("\tERROR: NO SHOPIFY ORDERS FOUND!");
                accountingItem.ErrorMessage = "Shopify: No orders found";
                accountingItem.NumPurchase = "NOT FOUND";
            }
        }

        static void ProcessAliExpressMatch(HashSet<string> usedOrderNumbers, IOrderedEnumerable<AliExpressOrderGroup> aliExpressQuery, List<OberloOrder> oberloOrders, AccountingItem accountingItem)
        {
            // flatten the aliexpress order list
            var aliExpressOrderList = aliExpressQuery.SelectMany(a => a.Children).ToList();

            // join the aliexpress list and the oberlo list on aliexpress order number
            var joined = from a in aliExpressOrderList
                         join b in oberloOrders
                        on a.OrderId.ToString() equals b.AliOrderNumber
                         select new { AliExpress = a, Oberlo = b };

            if (joined.Count() > 0)
            {
                Console.WriteLine("\tSHOPIFY ORDERS FOUND ...");

                string orderNumber = "NONE FOUND";
                foreach (var item in joined.Reverse())
                {
                    orderNumber = item.Oberlo.OrderNumber;
                    if (!usedOrderNumbers.Contains(orderNumber))
                    {
                        usedOrderNumbers.Add(orderNumber);
                        accountingItem.NumPurchase = orderNumber;
                        accountingItem.CustomerName = item.Oberlo.CustomerName;
                        Console.WriteLine("\tSELECTED: {0} {1}", accountingItem.NumPurchase, accountingItem.CustomerName);
                        break;
                    }
                }
            }

            // could not find shopify order numbers
            else
            {
                Console.WriteLine("\tERROR: NO OBERLO ORDERS FOUND!");
                var orderIds = string.Join(", ", Array.ConvertAll(aliExpressOrderList.ToArray(), i => i.OrderId));
                var orderCustomers = string.Join(", ", Array.ConvertAll(aliExpressOrderList.ToArray(), i => i.ContactName));
                accountingItem.ErrorMessage = string.Format("Oberlo: No shopify order found for order {0} ({1})", orderIds, orderCustomers);
                accountingItem.NumPurchase = "NOT FOUND";
            }
        }
        #endregion
    }
}
