using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Google.Apis.Sheets.v4.Data;

namespace AccountingServices.GoogleSheets
{
    public static class GoogleSheetsRequests
    {
        public static Request GetAppendRowsRequest(int sheetId, int numberOfRows)
        {
            Request appendRowsRequest = new Request()
            {
                AppendDimension = new AppendDimensionRequest()
                {
                    SheetId = sheetId,
                    Dimension = "ROWS",
                    Length = numberOfRows
                }
            };
            return appendRowsRequest;
        }

        public static Request GetAppendColumnsRequest(int sheetId, int numberOfColumns)
        {
            Request appendColumnsRequest = new Request()
            {
                AppendDimension = new AppendDimensionRequest()
                {
                    SheetId = sheetId,
                    Dimension = "COLUMNS",
                    Length = numberOfColumns
                }
            };
            return appendColumnsRequest;
        }

        public static Request GetDeleteRowsRequest(int sheetId, int rowStartIndex, int rowEndIndex)
        {
            Request deleteRowsRequest = new Request()
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
            return deleteRowsRequest;
        }

        public static Request GetDeleteColumnsRequest(int sheetId, int columnStartIndex, int columnEndIndex)
        {
            Request deleteColumnsRequest = new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = sheetId,
                        Dimension = "COLUMNS",
                        StartIndex = columnStartIndex,
                        EndIndex = columnEndIndex
                    }
                }
            };
            return deleteColumnsRequest;
        }

        public static Request GetInsertRowsRequest(int sheetId, int startRowIndex, int endRowIndex, bool doInheritFromBefore = false)
        {
            Request insertRowsRequest = new Request()
            {
                InsertDimension = new InsertDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = sheetId,
                        Dimension = "ROWS",
                        StartIndex = startRowIndex,
                        EndIndex = endRowIndex
                    },
                    InheritFromBefore = doInheritFromBefore
                }
            };

            return insertRowsRequest;
        }

        public static Request GetInsertColumnsRequest(int sheetId, int startColumnIndex, int endColumnIndex, bool doInheritFromBefore = false)
        {
            Request insertRowsRequest = new Request()
            {
                InsertDimension = new InsertDimensionRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = sheetId,
                        Dimension = "COLUMNS",
                        StartIndex = startColumnIndex,
                        EndIndex = endColumnIndex
                    },
                    InheritFromBefore = doInheritFromBefore
                }
            };

            return insertRowsRequest;
        }

        public static Request HideColumnsRequest(int sheetId, string startColumn, string endColumn)
        {
            int startColumnIndex = GoogleSheetsRequests.ColumnNumber(startColumn) - 1;
            int endColumnIndex = GoogleSheetsRequests.ColumnNumber(endColumn);

            return HideColumnsRequest(sheetId, startColumnIndex, endColumnIndex);
        }

        public static Request HideColumnsRequest(int sheetId, int startColumnIndex, int endColumnIndex)
        {
            Request hideColumnsRequest = new Request()
            {
                UpdateDimensionProperties = new UpdateDimensionPropertiesRequest()
                {
                    Range = new DimensionRange()
                    {
                        SheetId = sheetId,
                        Dimension = "COLUMNS",
                        StartIndex = startColumnIndex,
                        EndIndex = endColumnIndex
                    },
                    Properties = new DimensionProperties()
                    {
                        HiddenByUser = true,
                    },
                    Fields = "hiddenByUser"
                }
            };
            return hideColumnsRequest;
        }

        public static Request FreezeRowsRequest(int sheetId, int freezeRowCount)
        {
            Request freezeRowsRequest = new Request()
            {
                UpdateSheetProperties = new UpdateSheetPropertiesRequest()
                {
                    Properties = new SheetProperties()
                    {
                        SheetId = sheetId,
                        GridProperties = new GridProperties()
                        {
                            FrozenRowCount = freezeRowCount
                        }
                    },
                    Fields = "gridProperties.frozenRowCount"
                }
            };
            return freezeRowsRequest;
        }



        public static Request GetFindReplaceRequest(int sheetId, string query, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            var findReplaceRequest = new FindReplaceRequest()
            {
                Find = query,
                MatchCase = true,
                MatchEntireCell = true,
                SearchByRegex = false,
                IncludeFormulas = false,
                Range = new GridRange()
                {
                    SheetId = sheetId,
                    StartColumnIndex = startColumnIndex,
                    EndColumnIndex = endColumnIndex,
                    StartRowIndex = startRowIndex,
                    EndRowIndex = endRowIndex
                },
                SheetId = sheetId,
                AllSheets = false
            };
            var request = new Request() { FindReplace = findReplaceRequest };
            return request;
        }

        public static Request GetAppendCellsRequest(int sheetId, string[] columns, int fgColorHeader, int bgColorHeader, bool doFullFormatting)
        {
            var appendCellsRequestHeader = CreateAppendCellRequest(sheetId, columns, fgColorHeader, bgColorHeader, doFullFormatting);
            var request = new Request() { AppendCells = appendCellsRequestHeader };
            return request;
        }

        public static List<Request> GetAppendDataTableRequests(int sheetId, DataTable dt, bool doUseTableHeaders, bool doUseSubTotalsAtTop)
        {
            return GetAppendDataTableRequests(sheetId, dt, -1, -1, -1, -1, doUseTableHeaders, doUseSubTotalsAtTop);
        }

        public static List<Request> GetAppendDataTableRequests(int sheetId, DataTable dt, int fgColorHeader, int bgColorHeader, int fgColorRow, int bgColorRow, bool doUseTableHeaders, bool doUseSubTotalsAtTop)
        {
            var requests = new List<Request>();

            if (dt != null)
            {
                // append headers
                if (doUseTableHeaders)
                {
                    var appendCellsRequestHeader = CreateAppendCellRequest(sheetId, dt.Columns, fgColorHeader, bgColorHeader, true);
                    requests.Add(new Request() { AppendCells = appendCellsRequestHeader });
                }

                // insert empty subtotal row in next row
                if (doUseSubTotalsAtTop)
                {
                    // initialize empty string array
                    var emptyRow = Enumerable.Repeat<string>(string.Empty, dt.Columns.Count).ToArray();
                    var appendCellsRequestSubTotals = CreateAppendCellRequest(sheetId, emptyRow, fgColorHeader, bgColorHeader, true);
                    requests.Add(new Request() { AppendCells = appendCellsRequestSubTotals });
                }

                // append rows
                var appendCellsRequest = CreateAppendCellRequest(sheetId, dt.Rows, fgColorRow, bgColorRow);
                requests.Add(new Request() { AppendCells = appendCellsRequest });

                return requests;
            }

            return null;
        }

        public static List<Request> GetInsertDataTableRequests(int sheetId, DataTable dt, int startRowIndex, int startColumnIndex, bool doUseTableHeaders)
        {
            return GetInsertDataTableRequests(sheetId, dt, startRowIndex, startColumnIndex, -1, -1, -1, -1, doUseTableHeaders);
        }

        public static List<Request> GetInsertDataTableRequests(int sheetId, DataTable dt, int startRowIndex, int startColumnIndex, int fgColorHeader, int bgColorHeader, int fgColorRow, int bgColorRow, bool doUseTableHeaders)
        {
            var requests = new List<Request>();

            if (dt != null)
            {
                int rowCounter = 0;
                // insert headers
                if (doUseTableHeaders)
                {
                    rowCounter++;
                    var updateCellsRequestHeader = CreateUpdateCellsRequest(sheetId, startRowIndex, startColumnIndex, dt.Columns, fgColorHeader, bgColorHeader, true);
                    requests.Add(new Request() { UpdateCells = updateCellsRequestHeader });
                }

                // insert rows
                var updateCellsRequest = CreateUpdateCellsRequest(sheetId, startRowIndex + rowCounter, startColumnIndex, dt.Rows, fgColorRow, bgColorRow);
                requests.Add(new Request() { UpdateCells = updateCellsRequest });

                return requests;
            }

            return null;
        }

        public static Request GetAddSheetRequest(string sheetName, int columnCount)
        {
            // add new sheet
            var addSheetRequest = new Request()
            {
                AddSheet = new AddSheetRequest()
                {
                    Properties = new SheetProperties()
                    {
                        Title = sheetName,
                        GridProperties = new GridProperties()
                        {
                            ColumnCount = columnCount
                        }
                    }
                }
            };
            return addSheetRequest;
        }

        public static Request GetBasicFilterRequest(int sheetId, int startRowIndex, int endRowIndex, string startColumn, string endColumn)
        {
            int startColumnIndex = GoogleSheetsRequests.ColumnNumber(startColumn) - 1;
            int endColumnIndex = GoogleSheetsRequests.ColumnNumber(endColumn);

            return GetBasicFilterRequest(sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
        }

        public static Request GetBasicFilterRequest(int sheetId, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
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

        public static Request GetAutoResizeColumnsRequest(int sheetId, int startColumnIndex, int endColumnIndex)
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

        public static Request GetFormulaRequest(int sheetId, string formulaValue, int startRowIndex, int endRowIndex, string startColumn, string endColumn)
        {
            int startColumnIndex = GoogleSheetsRequests.ColumnNumber(startColumn) - 1;
            int endColumnIndex = GoogleSheetsRequests.ColumnNumber(endColumn);

            return GetFormulaRequest(sheetId, formulaValue, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
        }

        public static Request GetFormulaRequest(int sheetId, string formulaValue, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
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

        public static Request GetNumberFormatRequest(int sheetId, string numberFormatPattern, int fgColor, int bgColor, int startRowIndex, int endRowIndex, string startColumn, string endColumn)
        {
            int startColumnIndex = GoogleSheetsRequests.ColumnNumber(startColumn) - 1;
            int endColumnIndex = GoogleSheetsRequests.ColumnNumber(endColumn);

            return GetNumberFormatRequest(sheetId, numberFormatPattern, fgColor, bgColor, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
        }

        public static Request GetNumberFormatRequest(int sheetId, string numberFormatPattern, int fgColor, int bgColor, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            var numberFormat = new CellFormat()
            {
                NumberFormat = new NumberFormat()
                {
                    Type = "NUMBER",
                    Pattern = numberFormatPattern // e.g. "#,##0.00;[Red]-#,##0.00;"
                },
                BackgroundColor = GoogleSheetsRequests.GetColor(bgColor),
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsRequests.GetColor(fgColor)
                }
            };

            var numberFormatRequest = new Request()
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
                        UserEnteredFormat = numberFormat
                    },
                    Fields = "UserEnteredFormat"
                }
            };
            return numberFormatRequest;
        }

        public static Request GetFormulaAndNumberFormatRequest(int sheetId, string formulaValue, int startRowIndex, int endRowIndex, string startColumn, string endColumn)
        {
            int startColumnIndex = GoogleSheetsRequests.ColumnNumber(startColumn) - 1;
            int endColumnIndex = GoogleSheetsRequests.ColumnNumber(endColumn);

            return GetFormulaAndNumberFormatRequest(sheetId, formulaValue, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
        }

        public static Request GetFormulaAndNumberFormatRequest(int sheetId, string formulaValue, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            var numberFormat = new CellFormat()
            {
                NumberFormat = new NumberFormat()
                {
                    Type = "NUMBER",
                    Pattern = "#,##0.00;[Red]-#,##0.00;"
                }
            };

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
                            FormulaValue = formulaValue,
                        },
                        UserEnteredFormat = numberFormat
                    },
                    Fields = "UserEnteredValue,UserEnteredFormat"
                }
            };
            return formulaRequest;
        }

        public static Request GetFormulaAndTextFormatRequest(int sheetId, string formulaValue, int fgColor, int bgColor, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            var userEnteredFormat = new CellFormat()
            {
                BackgroundColor = GoogleSheetsRequests.GetColor(bgColor),
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsRequests.GetColor(fgColor),
                    FontSize = 11,
                    Bold = true
                }
            };

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
                            FormulaValue = formulaValue,
                        },
                        UserEnteredFormat = userEnteredFormat
                    },
                    Fields = "UserEnteredValue,UserEnteredFormat"
                }
            };
            return formulaRequest;
        }

        public static Request GetFormatRequest(int sheetId, int fgColor, int bgColor, int startRowIndex, int endRowIndex, string startColumn, string endColumn)
        {
            int startColumnIndex = GoogleSheetsRequests.ColumnNumber(startColumn) - 1;
            int endColumnIndex = GoogleSheetsRequests.ColumnNumber(endColumn);

            return GetFormatRequest(sheetId, fgColor, bgColor, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
        }

        public static Request GetFormatRequest(int sheetId, int fgColor, int bgColor, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            // define format
            var userEnteredFormat = new CellFormat()
            {
                BackgroundColor = GoogleSheetsRequests.GetColor(bgColor),
                //HorizontalAlignment = "LEFT",
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GoogleSheetsRequests.GetColor(fgColor),
                    FontSize = 11,
                    Bold = false
                }
            };

            // create the request
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
            return formatRequest;
        }

        #region Append Cell Request and Data Table Request
        private static AppendCellsRequest CreateAppendCellRequest(int sheetId, string[] columns, int fgColorHeader, int bgColorHeader, bool doFullFormatting)
        {
            var rowData = CreateRowData(sheetId, columns, fgColorHeader, bgColorHeader, doFullFormatting);

            var rowDataList = new List<RowData>();
            rowDataList.Add(rowData);

            var appendRequest = new AppendCellsRequest();
            appendRequest.SheetId = sheetId;
            appendRequest.Rows = rowDataList;
            appendRequest.Fields = "*";
            return appendRequest;
        }

        private static AppendCellsRequest CreateAppendCellRequest(int sheetId, DataColumnCollection columns, int fgColorHeader, int bgColorHeader, bool doFullFormatting)
        {
            var rowData = CreateRowData(sheetId, columns, fgColorHeader, bgColorHeader, doFullFormatting);

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

        public static Request GetUpdateCellsRequest(int sheetId, int startRowIndex, int startColumnIndex, string[] columns, int fgColorHeader, int bgColorHeader, bool doFullFormatting)
        {
            var updateCellsRequestHeader = CreateUpdateCellsRequest(sheetId, startRowIndex, startColumnIndex, columns, fgColorHeader, bgColorHeader, doFullFormatting);
            var request = new Request() { UpdateCells = updateCellsRequestHeader };
            return request;
        }

        private static UpdateCellsRequest CreateUpdateCellsRequest(int sheetId, int startRowIndex, int startColumnIndex, string[] columns, int fgColorHeader, int bgColorHeader, bool doFullFormatting)
        {
            var rowData = CreateRowData(sheetId, columns, fgColorHeader, bgColorHeader, doFullFormatting);

            var rowDataList = new List<RowData>();
            rowDataList.Add(rowData);

            var updateRequest = new UpdateCellsRequest();
            updateRequest.Start = new GridCoordinate()
            {
                SheetId = sheetId,
                ColumnIndex = startColumnIndex,
                RowIndex = startRowIndex,
            };
            updateRequest.Rows = rowDataList;
            updateRequest.Fields = "UserEnteredValue,UserEnteredFormat";
            return updateRequest;
        }

        private static UpdateCellsRequest CreateUpdateCellsRequest(int sheetId, int startRowIndex, int startColumnIndex, DataColumnCollection columns, int fgColorHeader, int bgColorHeader, bool doFullFormatting)
        {
            var rowData = CreateRowData(sheetId, columns, fgColorHeader, bgColorHeader, doFullFormatting);

            var rowDataList = new List<RowData>();
            rowDataList.Add(rowData);

            var updateRequest = new UpdateCellsRequest();
            updateRequest.Start = new GridCoordinate()
            {
                SheetId = sheetId,
                ColumnIndex = startColumnIndex,
                RowIndex = startRowIndex,
            };
            updateRequest.Rows = rowDataList;
            updateRequest.Fields = "UserEnteredValue,UserEnteredFormat";
            return updateRequest;
        }

        private static UpdateCellsRequest CreateUpdateCellsRequest(int sheetId, int startRowIndex, int startColumnIndex, DataRowCollection rows, int fgColorRow, int bgColorRow)
        {
            var rowDataList = new List<RowData>();
            foreach (DataRow row in rows)
            {
                var rowData = CreateRowData(sheetId, row, fgColorRow, bgColorRow);
                rowDataList.Add(rowData);
            }

            var updateRequest = new UpdateCellsRequest();
            updateRequest.Start = new GridCoordinate()
            {
                SheetId = sheetId,
                ColumnIndex = startColumnIndex,
                RowIndex = startRowIndex,
            };
            updateRequest.Rows = rowDataList;
            updateRequest.Fields = "UserEnteredValue,UserEnteredFormat";
            return updateRequest;
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
                    ForegroundColor = GetColor(fgColorRow)
                }
            };

            var dateFormat = new CellFormat()
            {
                BackgroundColor = GetColor(bgColorRow),
                TextFormat = new TextFormat()
                {
                    ForegroundColor = GetColor(fgColorRow)
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
                    ForegroundColor = GetColor(fgColorRow)
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
                    ForegroundColor = GetColor(fgColorRow)
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

        private static RowData CreateRowData(int sheetId, DataColumnCollection columns, int fgColorHeader, int bgColorHeader, bool doFullFormatting)
        {
            string[] columnNames = columns.Cast<DataColumn>()
                                             .Select(x => x.ColumnName)
                                             .ToArray();

            return CreateRowData(sheetId, columnNames, fgColorHeader, bgColorHeader, doFullFormatting);

        }

        private static RowData CreateRowData(int sheetId, string[] columns, int fgColorHeader, int bgColorHeader, bool doFullFormatting)
        {
            // https://github.com/opendatakit/aggregate/blob/master/src/main/java/org/opendatakit/aggregate/externalservice/GoogleSpreadsheet.java

            // define header cell format
            CellFormat headerFormat = null;

            if (doFullFormatting)
            {
                headerFormat = new CellFormat()
                {
                    BackgroundColor = GetColor(bgColorHeader),
                    HorizontalAlignment = "CENTER",
                    TextFormat = new TextFormat()
                    {
                        ForegroundColor = GetColor(fgColorHeader),
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
            }
            else
            {
                headerFormat = new CellFormat()
                {
                    BackgroundColor = GetColor(bgColorHeader),
                    TextFormat = new TextFormat()
                    {
                        ForegroundColor = GetColor(fgColorHeader),
                    },
                };
            }

            var cellDataList = new List<CellData>();
            foreach (var item in columns)
            {
                var cellData = new CellData();
                cellData.UserEnteredValue = new ExtendedValue();

                if (item != null)
                {
                    cellData.UserEnteredValue.StringValue = item;
                }
                cellData.UserEnteredFormat = headerFormat;
                cellDataList.Add(cellData);
            }

            var rowData = new RowData()
            {
                Values = cellDataList
            };

            return rowData;
        }
        #endregion

        public static Color GetColor(int argb)
        {
            if (argb < 0) return null;

            var c = System.Drawing.Color.FromArgb(argb);

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

        public static string ColumnAddress(int columnNumber)
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

        public static int ColumnNumber(string columnAddress)
        {
            int[] digits = new int[columnAddress.Length];
            for (int i = 0; i < columnAddress.Length; ++i)
            {
                digits[i] = Convert.ToInt32(columnAddress[i]) - 64;
            }
            int mul = 1; int res = 0;
            for (int pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }
            return res;
        }
    }
}