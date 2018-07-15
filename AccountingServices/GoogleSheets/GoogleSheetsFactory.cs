using System;
using System.IO;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Linq;
using Newtonsoft.Json;
using System.Data;

namespace AccountingServices.GoogleSheets
{
    public class GoogleSheetsFactory : IDisposable
    {
        public static readonly string[] SCOPES = { SheetsService.Scope.Spreadsheets };
        public static readonly string APPLICATION_NAME = "Wazalo Accounting";
        public static readonly string SPREADSHEET_ID = "1mGFDwqV0rb707hkdCEwytA5-JzWOC8dH3Keb6ipV8L8";
        public SheetsService Service { get; set; }

        public GoogleSheetsFactory()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("google_client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                  .CreateScoped(SCOPES);
            }

            // Create Google Sheets API service.
            this.Service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APPLICATION_NAME,
            });
        }

        public int GetSheetIdFromSheetName(string sheetName)
        {
            // get sheet id by sheet name
            var spreadsheet = Service.Spreadsheets.Get(SPREADSHEET_ID).Execute();
            var sheet = spreadsheet.Sheets.Where(s => s.Properties.Title == sheetName).FirstOrDefault();
            if (sheet != null && sheet.Properties != null)
            {
                return (int)sheet.Properties.SheetId;
            }
            return -1;
        }

        public Sheet GetSheetFromSheetName(string sheetName)
        {
            // get sheet id by sheet name
            var spreadsheet = Service.Spreadsheets.Get(SPREADSHEET_ID).Execute();
            var sheet = spreadsheet.Sheets.Where(s => s.Properties.Title == sheetName).FirstOrDefault();
            return sheet;
        }

        public int AddSheet(string sheetName, int columnCount = 26)
        {
            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();

            // add the add sheet request
            batchUpdateSpreadsheetRequest.Requests.Add(GoogleSheetsRequests.GetAddSheetRequest(sheetName, columnCount));

            var batchUpdateRequest = Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SPREADSHEET_ID);
            var response = batchUpdateRequest.Execute();
            if (response.Replies.Count() > 0)
            {
                AddSheetResponse addSheetResponse = (AddSheetResponse)response.Replies.FirstOrDefault().AddSheet;
                return addSheetResponse.Properties.SheetId.Value;
            }
            return -1;
        }

        public void AppendRow(string sheetName, string headerRange, IEnumerable<string> rowData)
        {
            var range = $"{sheetName}!{headerRange}";

            var valueRange = new ValueRange();

            //List<object> objectlist = rowData.Select(x => string.IsNullOrEmpty(x) ? "" : x).Cast<object>().ToList();
            List<object> objectlist = rowData.Cast<object>().ToList();
            valueRange.Values = new List<IList<object>> { objectlist };

            var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SPREADSHEET_ID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            var appendResponse = appendRequest.Execute();
            Console.WriteLine("AppendRow:\n" + JsonConvert.SerializeObject(appendResponse));
        }

        public void UpdateRow(string sheetName, string headerRange, IEnumerable<string> rowData)
        {
            var range = $"{sheetName}!{headerRange}";

            var valueRange = new ValueRange();

            //List<object> objectlist = rowData.Select(x => string.IsNullOrEmpty(x) ? "" : x).Cast<object>().ToList();
            List<object> objectlist = rowData.Cast<object>().ToList();
            valueRange.Values = new List<IList<object>> { objectlist };

            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, SPREADSHEET_ID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            Console.WriteLine("UpdateRow:\n" + JsonConvert.SerializeObject(updateResponse));
        }

        public DataTable ReadDataTable(string sheetName, int startRowNumber, int endRowNumber, bool doUseSubTotalsAtTop)
        {
            List<string> ranges = new List<string>();
            var fullRange = $"{sheetName}!{startRowNumber}:{endRowNumber}";
            ranges.Add(fullRange);

            // determine the data type for the first row below the header
            // https://stackoverflow.com/questions/46647135/how-to-determine-the-data-type-of-values-returned-by-google-sheets-api
            var request = Service.Spreadsheets.Get(SPREADSHEET_ID);
            request.Ranges = ranges;
            request.IncludeGridData = false;
            request.Fields = "sheets(data(rowData(values(userEnteredFormat/numberFormat,userEnteredValue)),startColumn,startRow))";
            var response = request.Execute();

            var rowData = response.Sheets.FirstOrDefault().Data.FirstOrDefault().RowData;

            // header
            var headerValues = rowData[0].Values;

            var headerList = new List<string>();
            foreach (var header in headerValues)
            {
                if (header.UserEnteredValue != null && header.UserEnteredValue.StringValue != null)
                {
                    headerList.Add(header.UserEnteredValue.StringValue);
                }
            }

            DataTable dt = new DataTable();
            int startRowIndex = 1;
            int endRowCount = rowData.Count;

            if (doUseSubTotalsAtTop)
            {
                startRowIndex++;
            }
            else
            {
                endRowCount--; // include all rows except the subtotal at the bottom
            }

            for (int i = startRowIndex; i < endRowCount; i++)
            {
                var rowValues = rowData[i].Values;

                if (i == startRowIndex) // first data row
                {
                    // read the first line of data and get data types
                    var dataTypeAndValueList = new List<KeyValuePair<Type, object>>();
                    foreach (var dataRow in rowValues)
                    {
                        var pair = GetDataTypeAndValueFromDataRow(dataRow);
                        dataTypeAndValueList.Add(pair);
                    }

                    // build data table columns
                    if (headerList.Count != dataTypeAndValueList.Count)
                    {
                        Console.WriteLine("Error! Failed reading datatable!");
                        return null;
                    }

                    // add columns
                    // add row number
                    dt.Columns.Add("RowNumber", typeof(int));

                    // add the rest of the headers
                    for (int j = 0; j < headerList.Count; j++)
                    {
                        dt.Columns.Add(headerList[j], dataTypeAndValueList[j].Key);
                    }

                    // add first row of values
                    DataRow workRow = dt.NewRow();

                    // add row number
                    workRow[0] = startRowNumber + i;

                    // add the rest of the values
                    for (int k = 0; k < headerList.Count; k++)
                    {
                        object value = dataTypeAndValueList[k].Value;
                        if (dataTypeAndValueList[k].Key == typeof(DateTime))
                        {
                            // Google Sheets uses a form of epoch date that is commonly used in spreadsheets. 
                            // The whole number portion of the value (left of the decimal) counts the days since 
                            // December 30th 1899. The fractional portion (right of the decimal) 
                            // counts the time as a fraction of one day. 
                            // For example, January 1st 1900 at noon would be 2.5, 
                            // 2 because it's two days after December 30th, 1899, 
                            // and .5 because noon is half a day. 
                            // February 1st 1900 at 3pm would be 33.625.
                            value = DateTime.FromOADate((double)value);
                        }
                        workRow[k + 1] = (value == null ? DBNull.Value : value);
                    }
                    dt.Rows.Add(workRow);
                }
                else
                {
                    // second and the rest of the data rows
                    // add first row of values

                    // read the first line of data and get data types
                    DataRow workRow = dt.NewRow();

                    // add row number
                    workRow[0] = startRowNumber + i;

                    // add the rest of the values
                    if (rowValues != null && rowValues.Count > 0)
                    {
                        for (int j = 0; j < rowValues.Count; j++)
                        {
                            var pair = GetDataTypeAndValueFromDataRow(rowValues[j]);

                            object value = pair.Value;
                            if (pair.Key == typeof(DateTime))
                            {
                                // Google Sheets uses a form of epoch date that is commonly used in spreadsheets. 
                                // The whole number portion of the value (left of the decimal) counts the days since 
                                // December 30th 1899. The fractional portion (right of the decimal) 
                                // counts the time as a fraction of one day. 
                                // For example, January 1st 1900 at noon would be 2.5, 
                                // 2 because it's two days after December 30th, 1899, 
                                // and .5 because noon is half a day. 
                                // February 1st 1900 at 3pm would be 33.625.
                                value = DateTime.FromOADate((double)value);
                            }
                            workRow[j + 1] = (value == null ? DBNull.Value : value);
                        }
                    }
                    dt.Rows.Add(workRow);
                }
            }

            return dt;
        }

        private static KeyValuePair<Type, object> GetDataTypeAndValueFromDataRow(CellData data)
        {
            Type type = null;
            object value = null;
            if (data.UserEnteredFormat != null)
            {
                if (data.UserEnteredFormat.NumberFormat != null)
                {
                    switch (data.UserEnteredFormat.NumberFormat.Type)
                    {
                        case "TEXT":
                            type = typeof(string);
                            break;
                        case "NUMBER":
                            type = typeof(decimal);
                            break;
                        case "DATE":
                            type = typeof(DateTime);
                            break;
                        default:
                            break;
                    }
                }
            }

            // no userEnteredFormat
            if (data.UserEnteredValue != null)
            {
                if (data.UserEnteredValue.NumberValue.HasValue)
                {
                    if (type == null) type = typeof(decimal);
                    value = data.UserEnteredValue.NumberValue.Value;
                }
                else if (data.UserEnteredValue.StringValue != null)
                {
                    if (type == null) type = typeof(string);
                    value = data.UserEnteredValue.StringValue;
                }
                else if (data.UserEnteredValue.BoolValue.HasValue)
                {
                    if (type == null) type = typeof(bool);
                    value = data.UserEnteredValue.BoolValue.Value;
                }
                else if (data.UserEnteredValue.FormulaValue != null)
                {
                    if (type == null) type = typeof(string);
                    if (type == typeof(decimal))
                    {
                        value = 0;
                    }
                    // always use string if there is a formula value
                    //type = typeof(string);
                    //value = data.UserEnteredValue.FormulaValue;
                }
                else if (data.UserEnteredValue.ErrorValue != null)
                {
                    // ignore?
                }
            }
            else
            {
                // UserEnteredValue is null
                type = typeof(string); // key cannot be null
                value = null;
            }
            return new KeyValuePair<Type, object>(type, value);
        }

        public void Dispose()
        {
            Service = null;
        }
    }
}