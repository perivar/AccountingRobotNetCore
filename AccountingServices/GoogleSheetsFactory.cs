using System;
using System.IO;
using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Linq;
using Newtonsoft.Json;

namespace AccountingServices
{
    public class GoogleSheetsFactory
    {
        static readonly string[] scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string applicationName = "Wazalo Accounting";
        static readonly string spreadsheetId = "1mGFDwqV0rb707hkdCEwytA5-JzWOC8dH3Keb6ipV8L8";
        static SheetsService service;

        public GoogleSheetsFactory()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("google_client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                  .CreateScoped(scopes);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
        }

        public int GetSheetIdFromSheetName(string sheetName)
        {
            // get sheet id by sheet name
            var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets.Where(s => s.Properties.Title == sheetName).FirstOrDefault();
            int sheetId = (int)sheet.Properties.SheetId;

            return sheetId;
        }

        public Sheet GetSheetFromSheetName(string sheetName)
        {
            // get sheet id by sheet name
            var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets.Where(s => s.Properties.Title == sheetName).FirstOrDefault();
            return sheet;
        }

        public int AddSheet(string sheetName)
        {
            // add new sheet
            var addSheetRequest = new AddSheetRequest();
            addSheetRequest.Properties = new SheetProperties();
            addSheetRequest.Properties.Title = sheetName;

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();
            batchUpdateSpreadsheetRequest.Requests.Add(new Request
            {
                AddSheet = addSheetRequest
            });

            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);

            var response = batchUpdateRequest.Execute();
            if (response.Replies.Count() > 0)
            {
                AddSheetResponse addSheetResponse = (AddSheetResponse)response.Replies.FirstOrDefault().AddSheet;
                return addSheetResponse.Properties.SheetId.Value;
            }

            return -1;
        }

        public void AddContentToSheet(int sheetId)
        {
            string[] colNames = new[] { "timestamp", "videoid", "videoname", "firstname", "lastname", "email" };

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();

            // Create starting coordinate where data would be written to
            var gridCoordinate = new GridCoordinate();
            gridCoordinate.ColumnIndex = 0;
            gridCoordinate.RowIndex = 0;
            gridCoordinate.SheetId = sheetId;

            var request = new Request();
            request.UpdateCells = new UpdateCellsRequest();

            batchUpdateSpreadsheetRequest.Requests.Add(request); // It's a batch request so you can create more than one request and send them all in one batch. Just use reqs.Requests.Add() to add additional requests for the same spreadsheet

            request.UpdateCells.Start = gridCoordinate;
            request.UpdateCells.Fields = "*"; // needed by API, throws error if null
            request.UpdateCells.Rows = new List<RowData>();

            //  add row
            var rd = new RowData();
            request.UpdateCells.Rows.Add(rd);
            rd.Values = new List<CellData>();

            // Assigning data to cells
            foreach (String column in colNames)
            {
                var cellData = new CellData();
                cellData.UserEnteredValue = new ExtendedValue();
                cellData.UserEnteredValue.StringValue = column;
                rd.Values.Add(cellData);
            }

            // Execute request
            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
            batchUpdateRequest.Execute();
        }

        public void UpdateFormatting(int sheetId, int color)
        {
            // https://developers.google.com/sheets/api/samples/formatting

            // define cell color
            var userEnteredFormat = new CellFormat()
            {
                BackgroundColor = GetColor(color),
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

            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
            batchUpdateRequest.Execute();
        }

        public void DeleteRows(int sheetId, int rowStartIndex, int rowEndIndex)
        {
            Request requestBody = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = sheetId,
                        Dimension = "ROWS",
                        StartIndex = rowStartIndex,
                        EndIndex = rowEndIndex
                    }
                }
            };

            List<Request> requests = new List<Request>();
            requests.Add(requestBody);

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = requests;

            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            Console.WriteLine(JsonConvert.SerializeObject(batchUpdateResponse));
        }

        public void InsertDataTable(string sheetName)
        {
            var range = $"{sheetName}!A:BA";

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

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            Console.WriteLine(JsonConvert.SerializeObject(updateResponse));
        }

        public void ReadEntries(string sheetName)
        {
            var range = $"{sheetName}!A:BA";
            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

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

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }

        public void UpdateEntry(string sheetName)
        {
            var range = $"{sheetName}!D543";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { "updated" };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }

        public void DeleteEntry(string sheetName)
        {
            var range = $"{sheetName}!A543:F";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
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

        private static Color GetColor(int argb)
        {
            System.Drawing.Color c = System.Drawing.Color.FromArgb(argb);

            // convert to float values
            var c1 = new Color()
            {
                Blue = (float)(c.B / 255.0f),
                Red = (float)(c.R / 255.0f),
                Green = (float)(c.G / 255.0f),
                Alpha = (float)(c.A / 255.0f)
            };

            return c1;
        }

    }
}