using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4.Data;

namespace AccountingServices
{
    public class GoogleSheetsBatchUpdateRequests : IDisposable
    {
        GoogleSheetsFactory factory;
        BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest;

        public GoogleSheetsBatchUpdateRequests()
        {
            this.factory = new GoogleSheetsFactory();
            this.batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
            this.batchUpdateSpreadsheetRequest.Requests = new List<Request>();
        }

        public void Add(Request request)
        {
            this.batchUpdateSpreadsheetRequest.Requests.Add(request);
        }

        public void Add(IList<Request> requests)
        {
            foreach (var request in requests)
            {
                this.batchUpdateSpreadsheetRequest.Requests.Add(request);
            }
        }

        public BatchUpdateSpreadsheetResponse Execute()
        {
            var batchUpdateRequest = factory.Service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, GoogleSheetsFactory.SPREADSHEET_ID);
            var batchUpdateResponse = batchUpdateRequest.Execute();
            return batchUpdateResponse;
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    factory = null;
                    batchUpdateSpreadsheetRequest = null;
                }

                disposedValue = true;
            }
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
        }
        #endregion
    }
}