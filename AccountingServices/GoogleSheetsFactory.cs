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

namespace AccountingServices
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

            //List<object> oblist = rowData.Select(x => string.IsNullOrEmpty(x) ? "" : x).Cast<object>().ToList();
            List<object> oblist = rowData.Cast<object>().ToList();
            valueRange.Values = new List<IList<object>> { oblist };

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

            //List<object> oblist = rowData.Select(x => string.IsNullOrEmpty(x) ? "" : x).Cast<object>().ToList();
            List<object> oblist = rowData.Cast<object>().ToList();
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, SPREADSHEET_ID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            Console.WriteLine("UpdateRow:\n" + JsonConvert.SerializeObject(updateResponse));
        }

        public DataTable ReadDataTable(string sheetName, int startRowNumber, int endRowNumber = 10000)
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
            for (int i = 0; i < rowData.Count - 1; i++) // include all rows except the subtotal at the bottom
            {
                if (i == 0) continue; // ship header, already been processed

                var rowValues = rowData[i].Values;
                if (i == 1) // first data row
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

        #region Methods for testing

        public void AppendDataTable(string sheetName, int sheetId, DataTable dt, int fgColorHeader, int bgColorHeader, int fgColorRow, int bgColorRow)
        {
            var range = $"{sheetName}!A:A";

            int startColumnIndex = 0;
            int endColumnIndex = dt.Columns.Count + 1;
            int startRowIndex = 1;
            int endRowIndex = dt.Rows.Count + 1;

            IList<IList<Object>> values = new List<IList<Object>>();
            if (dt != null)
            {
                // first add column names
                // and build subtotal list
                List<object> columnHeaders = new List<object>();
                List<object> subTotalFooters = new List<object>();
                int columnNumber = 1;
                foreach (DataColumn column in dt.Columns)
                {
                    string columnName = column.ColumnName;
                    columnHeaders.Add(columnName);

                    // =SUBTOTAL(109;O3:O174) = sum and ignore hidden values
                    subTotalFooters.Add(string.Format("=SUBTOTAL(109;{0}{1}:{0}{2})", GoogleSheetsRequests.ColumnAddress(columnNumber), startRowIndex + 2, endRowIndex + 1));

                    columnNumber++;
                }
                values.Add(columnHeaders);

                // then add row values
                foreach (DataRow row in dt.Rows)
                {
                    List<object> rowValues = row.ItemArray.ToList();
                    values.Add(rowValues);
                }

                // finally add the subtotal row
                values.Add(subTotalFooters);
            }

            ValueRange valueRange = new ValueRange() { Values = values };

            var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SPREADSHEET_ID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            var appendResponse = appendRequest.Execute();
            Console.WriteLine("AppendDataTable:\n" + JsonConvert.SerializeObject(appendResponse));

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();

            // define header cell format
            var userEnteredFormatHeader = new CellFormat()
            {
                BackgroundColor = GoogleSheetsRequests.GetColor(bgColorHeader),
                HorizontalAlignment = "CENTER",
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsRequests.GetColor(fgColorHeader),
                    FontSize = 11,
                    Bold = true
                },
                Borders = new Borders()
                {
                    Bottom = new Border()
                    {
                        Style = "DASHED",
                        Width = 2
                    },
                    Top = new Border()
                    {
                        Style = "DASHED",
                        Width = 2
                    }
                }
            };

            // create the update request for cells from the header row
            var formatRequestHeader = new Request()
            {
                RepeatCell = new RepeatCellRequest()
                {
                    Range = new GridRange()
                    {
                        SheetId = sheetId,
                        StartColumnIndex = startColumnIndex,
                        EndColumnIndex = endColumnIndex,
                        StartRowIndex = startRowIndex,
                        EndRowIndex = startRowIndex + 1 // only header
                    },
                    Cell = new CellData()
                    {
                        UserEnteredFormat = userEnteredFormatHeader
                    },
                    Fields = "UserEnteredFormat(BackgroundColor,TextFormat,HorizontalAlignment,Borders)"
                }
            };
            batchUpdateSpreadsheetRequest.Requests.Add(formatRequestHeader);


            // define row cell format
            var userEnteredFormatRows = new CellFormat()
            {
                BackgroundColor = GoogleSheetsRequests.GetColor(bgColorRow),
                HorizontalAlignment = "LEFT",
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsRequests.GetColor(fgColorRow),
                    FontSize = 11,
                    Bold = false
                }
            };

            // create the update request for cells from the header row
            var formatRequestRows = new Request()
            {
                RepeatCell = new RepeatCellRequest()
                {
                    Range = new GridRange()
                    {
                        SheetId = sheetId,
                        StartColumnIndex = startColumnIndex,
                        EndColumnIndex = endColumnIndex,
                        StartRowIndex = startRowIndex + 1,
                        EndRowIndex = endRowIndex + 1
                    },
                    Cell = new CellData()
                    {
                        UserEnteredFormat = userEnteredFormatRows
                    },
                    Fields = "UserEnteredFormat(BackgroundColor,TextFormat,HorizontalAlignment,Borders)"
                }
            };
            batchUpdateSpreadsheetRequest.Requests.Add(formatRequestRows);


            // set basic filter for all rows except the last
            var filterRequest = new Request()
            {
                SetBasicFilter = new SetBasicFilterRequest()
                {
                    Filter = new BasicFilter()
                    {
                        Criteria = null,
                        SortSpecs = null,
                        Range = new GridRange()
                        {
                            SheetId = sheetId,
                            StartColumnIndex = startColumnIndex,
                            EndColumnIndex = endColumnIndex,
                            StartRowIndex = startRowIndex,
                            EndRowIndex = endRowIndex + 1
                        }
                    }
                }
            };
            batchUpdateSpreadsheetRequest.Requests.Add(filterRequest);


            // auto resize the columns
            var autoResizeRequest = new Request()
            {
                AutoResizeDimensions = new AutoResizeDimensionsRequest()
                {
                    Dimensions = new DimensionRange()
                    {
                        SheetId = sheetId,
                        Dimension = "COLUMNS",
                        StartIndex = startColumnIndex,
                        EndIndex = endColumnIndex
                    }
                }
            };
            batchUpdateSpreadsheetRequest.Requests.Add(autoResizeRequest);

            var batchUpdateRequest = Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SPREADSHEET_ID);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            Console.WriteLine("AppendDataTable-Formatting:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
        }

        public void UpdateFormatting(int sheetId, CellFormat userEnteredFormat, int endColumnIndex, int endRowIndex, int startColumnIndex = 0, int startRowIndex = 0)
        {
            // https://developers.google.com/sheets/api/samples/formatting

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();

            // create the update format request for cells matching the grid range
            var formatRequest = new Request()
            {
                RepeatCell = new RepeatCellRequest()
                {
                    Range = new GridRange()
                    {
                        SheetId = sheetId,
                        StartColumnIndex = startColumnIndex,
                        EndColumnIndex = endColumnIndex,
                        StartRowIndex = startRowIndex,
                        EndRowIndex = endRowIndex
                    },
                    Cell = new CellData()
                    {
                        UserEnteredFormat = userEnteredFormat
                    },
                    Fields = "UserEnteredFormat(BackgroundColor,TextFormat,HorizontalAlignment,Borders)"
                }
            };
            batchUpdateSpreadsheetRequest.Requests.Add(formatRequest);

            var batchUpdateRequest = Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SPREADSHEET_ID);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            Console.WriteLine("UpdateFormatting:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
        }

        public void UpdateFormatting(int sheetId, int color)
        {
            // https://developers.google.com/sheets/api/samples/formatting

            // define cell color
            var userEnteredFormat = new CellFormat()
            {
                BackgroundColor = GoogleSheetsRequests.GetColor(color),
                TextFormat = new TextFormat()
                {
                    Bold = true
                },
                HorizontalAlignment = "CENTER",
                Borders = new Borders()
                {
                    Bottom = new Border()
                    {
                        Style = "DASHED",
                        Width = 2
                    },
                    Top = new Border()
                    {
                        Style = "DASHED",
                        Width = 2
                    }
                }
            };

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();

            // create the update request for cells from the first row
            var formatRequest = new Request()
            {
                RepeatCell = new RepeatCellRequest()
                {
                    Range = new GridRange()
                    {
                        SheetId = sheetId,
                        StartColumnIndex = 0,
                        EndColumnIndex = 4,
                        StartRowIndex = 0,
                        EndRowIndex = 1
                    },
                    Cell = new CellData()
                    {
                        UserEnteredFormat = userEnteredFormat
                    },
                    Fields = "UserEnteredFormat(BackgroundColor,TextFormat,HorizontalAlignment,Borders)"
                }
            };
            batchUpdateSpreadsheetRequest.Requests.Add(formatRequest);

            // set basic filter for all rows except the last
            var filterRequest = new Request()
            {
                SetBasicFilter = new SetBasicFilterRequest()
                {
                    Filter = new BasicFilter()
                    {
                        Criteria = null,
                        SortSpecs = null,
                        Range = new GridRange()
                        {
                            SheetId = sheetId,
                            StartColumnIndex = 0,
                            EndColumnIndex = 4,
                            StartRowIndex = 0,
                            EndRowIndex = 4
                        }
                    }
                }
            };
            batchUpdateSpreadsheetRequest.Requests.Add(filterRequest);

            /*
            FilterCriteria criteria = new FilterCriteria();
            criteria.Condition = new BooleanCondition();
            criteria.Condition.Type = "NOT_BLANK";

            var criteriaDictionary = new Dictionary<string, FilterCriteria>();
            criteriaDictionary.Add("8", criteria); // define at which index the  filter is active

            var filterRequest = new Request();
            filterRequest.AddFilterView = new AddFilterViewRequest();
            filterRequest.AddFilterView.Filter = new FilterView();
            filterRequest.AddFilterView.Filter.FilterViewId = 0;
            filterRequest.AddFilterView.Filter.Title = "Hide rows with errors";
            filterRequest.AddFilterView.Filter.Range = range1;
            filterRequest.AddFilterView.Filter.Criteria = criteriaDictionary;
            batchUpdateSpreadsheetRequest.Requests.Add(request);
             */

            var batchUpdateRequest = Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SPREADSHEET_ID);
            batchUpdateRequest.Execute();
        }

        public void BatchValuesUpdate(string sheetName)
        {
            // BatchUpdateValuesRequest is used to update several value ranges in one go
            var body = new BatchUpdateValuesRequest();
            body.ValueInputOption = ((int)SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED).ToString();

            List<ValueRange> valueRanges = new List<ValueRange>();
            ValueRange valueRange = new ValueRange();
            valueRange.Range = $"{sheetName}!A2:E2";

            IList<IList<object>> values = new List<IList<object>>();
            List<object> child = new List<object>();
            for (int i = 0; i < 5; i++)
            {
                child.Add(i);
            }
            values.Add(child);
            valueRange.Values = values;
            valueRanges.Add(valueRange);
            body.Data = valueRanges;

            var batchUpdateRequest = Service.Spreadsheets.Values.BatchUpdate(body, SPREADSHEET_ID);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            Console.WriteLine("BatchUpdate:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
        }

        public void InsertDataTest(string sheetName)
        {
            var range = $"{sheetName}!A:A";

            var formula1 = "=SUM(B2:B4)";
            var formula2 = "=SUM(C2:C4)";
            var formula3 = "=MAX(D2:D4)";

            List<object> list1 = new List<object>() { "Item", "Cost", "Stocked", "Ship Date" };
            List<object> list2 = new List<object>() { "Wheel", "$20,50", "4", "1/3/2016" };
            List<object> list3 = new List<object>() { "Door", "$15", "2", "15/3/2016" };
            List<object> list4 = new List<object>() { "Engine", "$100", "1", "20/12/2016" };
            List<object> list5 = new List<object>() { "Totals", formula1, formula2, formula3 };
            IList<IList<Object>> values = new List<IList<Object>>() { list1, list2, list3, list4, list5 };

            ValueRange valueRange = new ValueRange() { Values = values };

            /*
            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            var appendResponse = appendRequest.Execute();
            Console.WriteLine(JsonConvert.SerializeObject(appendResponse));
             */

            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, SPREADSHEET_ID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            Console.WriteLine(JsonConvert.SerializeObject(updateResponse));
        }

        public void ReadEntries(string sheetName)
        {
            var range = $"{sheetName}!A:BA";
            var request = Service.Spreadsheets.Values.Get(SPREADSHEET_ID, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    // Print columns A to F, which correspond to indices 0 and 4.
                    Console.WriteLine("{0} | {1} | {2} | {3} | {4} | {5}", row[0], row[1], row[2], row[3], row[4], row[5]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
        }
        public void CreateEntry(string sheetName)
        {
            var range = $"{sheetName}!A:F";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { "Hello!", "This", "was", "insertd", "via", "C#" };
            valueRange.Values = new List<IList<object>> { oblist };

            var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SPREADSHEET_ID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }

        public void UpdateEntry(string sheetName)
        {
            var range = $"{sheetName}!D543";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { "updated" };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = Service.Spreadsheets.Values.Update(valueRange, SPREADSHEET_ID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }

        public void DeleteEntry(string sheetName)
        {
            var range = $"{sheetName}!A543:F";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = Service.Spreadsheets.Values.Clear(requestBody, SPREADSHEET_ID, range);
            var deleteResponse = deleteRequest.Execute();
        }

        private static void UpdateGoogleSheetInBatch(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            SpreadsheetsResource.ValuesResource.AppendRequest request =
               service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = request.Execute();
        }
        #endregion
    }
}