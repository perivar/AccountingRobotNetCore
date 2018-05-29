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

        public void AppendRow(string sheetName, string headerRange, IEnumerable<string> rowData)
        {
            var range = $"{sheetName}!{headerRange}";

            var valueRange = new ValueRange();

            //List<object> oblist = rowData.Select(x => string.IsNullOrEmpty(x) ? "" : x).Cast<object>().ToList();
            List<object> oblist = rowData.Cast<object>().ToList();
            valueRange.Values = new List<IList<object>> { oblist };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
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

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
            Console.WriteLine("UpdateRow:\n" + JsonConvert.SerializeObject(updateResponse));
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

            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            Console.WriteLine("UpdateFormatting:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
        }

        public void AppendColumns(int sheetId, int numberOfColumns)
        {
            Request requestBody = new Request()
            {
                AppendDimension = new AppendDimensionRequest()
                {
                    SheetId = sheetId,
                    Dimension = "COLUMNS",
                    Length = numberOfColumns
                }
            };

            List<Request> requests = new List<Request>();
            requests.Add(requestBody);

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = requests;

            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            Console.WriteLine("AppendColumns:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
        }

        public void AppendRows(int sheetId, int numberOfRows)
        {
            Request requestBody = new Request()
            {
                AppendDimension = new AppendDimensionRequest()
                {
                    SheetId = sheetId,
                    Dimension = "ROWS",
                    Length = numberOfRows
                }
            };

            List<Request> requests = new List<Request>();
            requests.Add(requestBody);

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = requests;

            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            Console.WriteLine("AppendRows:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
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
            Console.WriteLine("DeleteRows:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
        }

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
                    subTotalFooters.Add(string.Format("=SUBTOTAL(109;{0}{1}:{0}{2})", GetExcelColumnName(columnNumber), startRowIndex + 2, endRowIndex + 1));

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

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            var appendResponse = appendRequest.Execute();
            Console.WriteLine("AppendDataTable:\n" + JsonConvert.SerializeObject(appendResponse));

            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();

            // define header cell format
            var userEnteredFormatHeader = new CellFormat()
            {
                BackgroundColor = GetColor(bgColorHeader),
                HorizontalAlignment = "CENTER",
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsFactory.GetColor(fgColorHeader),
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
                BackgroundColor = GetColor(bgColorRow),
                HorizontalAlignment = "LEFT",
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsFactory.GetColor(fgColorRow),
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

            var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            Console.WriteLine("AppendDataTable-Formatting:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
        }

        public void AppendDataTable(int sheetId, DataTable dt, int fgColorHeader, int bgColorHeader, int fgColorRow, int bgColorRow, bool autoResizeColumns = true)
        {
            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateSpreadsheetRequest.Requests = new List<Request>();

            if (dt != null)
            {
                int startColumnIndex = 0;
                int endColumnIndex = dt.Columns.Count + 1;
                int startRowIndex = 0;
                int endRowIndex = dt.Rows.Count + 1;

                // append headers
                var appendCellsRequestHeader = CreateAppendCellRequest(sheetId, dt.Columns, fgColorHeader, bgColorHeader);
                batchUpdateSpreadsheetRequest.Requests.Add(new Request() { AppendCells = appendCellsRequestHeader });

                // append rows
                var appendCellsRequest = CreateAppendCellRequest(sheetId, dt.Rows, fgColorRow, bgColorRow);
                batchUpdateSpreadsheetRequest.Requests.Add(new Request() { AppendCells = appendCellsRequest });

                // set basic filter for all rows
                batchUpdateSpreadsheetRequest.Requests.Add(
                    GetBasicFilterRequest(sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex)
                );

                // insert formula in column 1
                batchUpdateSpreadsheetRequest.Requests.Add(
                    GetFormulaRequest(sheetId,
                    string.Format("=IF(BA{0}=0;\" \";\"!!FEIL!!\")", startRowIndex + 2),
                    startRowIndex + 1, endRowIndex + 1, 0, 1)
                );

                if (autoResizeColumns)
                {
                    // auto resize columns
                    batchUpdateSpreadsheetRequest.Requests.Add(
                        GetAutoResizeColumnsRequest(sheetId, startColumnIndex, endColumnIndex)
                    );
                }

                var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
                var batchUpdateResponse = batchUpdateRequest.Execute();
                Console.WriteLine("AppendDataTable:\n" + JsonConvert.SerializeObject(batchUpdateResponse));
            }
        }

        private static Request GetBasicFilterRequest(int sheetId, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
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
                            EndRowIndex = endRowIndex
                        }
                    }
                }
            };
            return filterRequest;
        }

        private static Request GetAutoResizeColumnsRequest(int sheetId, int startColumnIndex, int endColumnIndex)
        {
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
            return autoResizeRequest;
        }

        private static Request GetFormulaRequest(int sheetId, string formulaValue, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            var formulaRequest = new Request()
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
                        UserEnteredValue = new ExtendedValue()
                        {
                            FormulaValue = formulaValue
                        }
                    },
                    Fields = "UserEnteredValue"
                }
            };
            return formulaRequest;
        }

        private static AppendCellsRequest CreateAppendCellRequest(int sheetId, DataColumnCollection columns, int fgColorHeader, int bgColorHeader)
        {
            var rowData = CreateRowData(sheetId, columns, fgColorHeader, bgColorHeader);

            var rowDataList = new List<RowData>();
            rowDataList.Add(rowData);

            var appendRequest = new AppendCellsRequest();
            appendRequest.SheetId = sheetId;
            appendRequest.Rows = rowDataList;
            appendRequest.Fields = "*";
            return appendRequest;
        }

        private static AppendCellsRequest CreateAppendCellRequest(int sheetId, DataRowCollection rows, int fgColorRow, int bgColorRow)
        {
            var rowDataList = new List<RowData>();
            foreach (DataRow row in rows)
            {
                var rowData = CreateRowData(sheetId, row, fgColorRow, bgColorRow);
                rowDataList.Add(rowData);
            }

            var appendRequest = new AppendCellsRequest();
            appendRequest.SheetId = sheetId;
            appendRequest.Rows = rowDataList;
            appendRequest.Fields = "*";
            return appendRequest;
        }

        private static RowData CreateRowData(int sheetId, DataRow row, int fgColorRow, int bgColorRow)
        {
            // https://github.com/opendatakit/aggregate/blob/master/src/main/java/org/opendatakit/aggregate/externalservice/GoogleSpreadsheet.java

            // define row cell formats
            var stringFormat = new CellFormat()
            {
                BackgroundColor = GetColor(bgColorRow),
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsFactory.GetColor(fgColorRow)
                }
            };

            var dateFormat = new CellFormat()
            {
                BackgroundColor = GetColor(bgColorRow),
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsFactory.GetColor(fgColorRow)
                },
                NumberFormat = new NumberFormat()
                {
                    Type = "DATE",
                    Pattern = "dd.MM.yyyy"
                }
            };

            var percentFormat = new CellFormat()
            {
                BackgroundColor = GetColor(bgColorRow),
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsFactory.GetColor(fgColorRow)
                },
                NumberFormat = new NumberFormat()
                {
                    Type = "NUMBER",
                    Pattern = "##.#%"
                }
            };

            var numberFormat = new CellFormat()
            {
                BackgroundColor = GetColor(bgColorRow),
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsFactory.GetColor(fgColorRow)
                },
                NumberFormat = new NumberFormat()
                {
                    Type = "NUMBER",
                    Pattern = "#,##0.00;[Red]-#,##0.00;"
                }
            };

            var cellDataList = new List<CellData>();
            foreach (var item in row.ItemArray)
            {
                var cellData = new CellData();

                if (item == null)
                {
                    cellData.UserEnteredValue = new ExtendedValue();
                }
                else
                {
                    var extendedValue = new ExtendedValue();
                    switch (item)
                    {
                        case bool boolValue:
                            extendedValue.BoolValue = boolValue;
                            cellData.UserEnteredValue = extendedValue;
                            cellData.UserEnteredFormat = stringFormat;
                            break;
                        case int intValue:
                            extendedValue.NumberValue = intValue;
                            cellData.UserEnteredValue = extendedValue;
                            cellData.UserEnteredFormat = numberFormat;
                            break;
                        case decimal decimalValue:
                            extendedValue.NumberValue = (double)decimalValue;
                            cellData.UserEnteredValue = extendedValue;
                            cellData.UserEnteredFormat = numberFormat;
                            break;
                        case DateTime dateTimeValue:
                            // 04.05.2018  23:59:00
                            // Google Sheets uses a form of epoch date that is commonly used in spreadsheets. 
                            // The whole number portion of the value (left of the decimal) counts the days since 
                            // December 30th 1899. The fractional portion (right of the decimal) 
                            // counts the time as a fraction of one day. 
                            // For example, January 1st 1900 at noon would be 2.5, 
                            // 2 because it's two days after December 30th, 1899, 
                            // and .5 because noon is half a day. 
                            // February 1st 1900 at 3pm would be 33.625.
                            extendedValue.NumberValue = dateTimeValue.ToOADate();
                            cellData.UserEnteredValue = extendedValue;
                            cellData.UserEnteredFormat = dateFormat;
                            break;
                        case string stringValue:
                            extendedValue.StringValue = stringValue;
                            cellData.UserEnteredValue = extendedValue;
                            cellData.UserEnteredFormat = stringFormat;
                            break;
                        default:
                            extendedValue.StringValue = item.ToString();
                            cellData.UserEnteredValue = extendedValue;
                            cellData.UserEnteredFormat = stringFormat;
                            break;
                    }
                }

                cellDataList.Add(cellData);
            }

            var rowData = new RowData()
            {
                Values = cellDataList
            };

            return rowData;
        }

        private static RowData CreateRowData(int sheetId, DataColumnCollection columns, int fgColorHeader, int bgColorHeader)
        {
            // https://github.com/opendatakit/aggregate/blob/master/src/main/java/org/opendatakit/aggregate/externalservice/GoogleSpreadsheet.java

            // define header cell format
            var headerFormat = new CellFormat()
            {
                BackgroundColor = GetColor(bgColorHeader),
                HorizontalAlignment = "CENTER",
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsFactory.GetColor(fgColorHeader),
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

            var cellDataList = new List<CellData>();
            foreach (DataColumn item in columns)
            {
                var cellData = new CellData();
                cellData.UserEnteredValue = new ExtendedValue();

                if (item != null)
                {
                    cellData.UserEnteredValue.StringValue = item.ColumnName;
                    cellData.UserEnteredFormat = headerFormat;
                }
                cellDataList.Add(cellData);
            }

            var rowData = new RowData()
            {
                Values = cellDataList
            };

            return rowData;
        }

        public static Color GetColor(int argb)
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

        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public void Dispose()
        {
            service = null;
        }


        #region Methods for testing
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

            var batchUpdateRequest = service.Spreadsheets.Values.BatchUpdate(body, spreadsheetId);
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
        #endregion
    }
}