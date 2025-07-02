using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetHelper
{
    public class GoogleSheetService
    {
        private readonly SheetsService _sheetsService;
        private readonly string _spreadsheetId;

        /// <summary>
        /// Initialize the Google Sheets service with credentials from a JSON file
        /// </summary>
        /// <param name="credentialsFilePath">Path to the JSON credentials file</param>
        /// <param name="spreadsheetId">The ID of the spreadsheet (found in the URL)</param>
        public GoogleSheetService(string credentialsFilePath, string spreadsheetId)
        {
            _spreadsheetId = spreadsheetId;
            
            using var stream = new FileStream(credentialsFilePath, FileMode.Open, FileAccess.Read);
            var credential = GoogleCredential.FromStream(stream)
                .CreateScoped(SheetsService.Scope.Spreadsheets);

            _sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Google Sheets API .NET 9 Client"
            });
        }

        /// <summary>
        /// Read data from a specific range in the spreadsheet
        /// </summary>
        /// <param name="range">The A1 notation of the range (e.g., "Sheet1!A1:D10")</param>
        /// <returns>A list of rows, where each row is a list of cell values</returns>
        public async Task<IList<IList<object>>> ReadAsync(string range)
        {
            var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = await request.ExecuteAsync();
            return response.Values ?? new List<IList<object>>();
        }

        /// <summary>
        /// Update data in a specific range in the spreadsheet
        /// </summary>
        /// <param name="range">The A1 notation of the range (e.g., "Sheet1!A1:D10")</param>
        /// <param name="values">The values to update</param>
        /// <returns>The number of updated cells</returns>
        public async Task<int> UpdateAsync(string range, IList<IList<object>> values)
        {
            var valueRange = new ValueRange { Values = values };
            var request = _sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            
            var response = await request.ExecuteAsync();
            return response.UpdatedCells ?? 0;
        }

        /// <summary>
        /// Append data to a specific range in the spreadsheet
        /// </summary>
        /// <param name="range">The A1 notation of the range (e.g., "Sheet1!A:D")</param>
        /// <param name="values">The values to append</param>
        /// <returns>The updated range</returns>
        public async Task<string> AppendAsync(string range, IList<IList<object>> values)
        {
            var valueRange = new ValueRange { Values = values };
            var request = _sheetsService.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            
            var response = await request.ExecuteAsync();
            return response.Updates?.UpdatedRange ?? string.Empty;
        }

        /// <summary>
        /// Clear data in a specific range in the spreadsheet
        /// </summary>
        /// <param name="range">The A1 notation of the range (e.g., "Sheet1!A1:D10")</param>
        /// <returns>The cleared range</returns>
        public async Task<string> ClearAsync(string range)
        {
            var request = _sheetsService.Spreadsheets.Values.Clear(new ClearValuesRequest(), _spreadsheetId, range);
            var response = await request.ExecuteAsync();
            return response.ClearedRange ?? string.Empty;
        }

        /// <summary>
        /// Delete rows from a spreadsheet
        /// </summary>
        /// <param name="sheetId">The sheet ID (not the spreadsheet ID)</param>
        /// <param name="startRowIndex">0-based index of the first row to delete</param>
        /// <param name="endRowIndex">0-based index of the last row to delete + 1</param>
        /// <returns>True if successful</returns>
        public async Task<bool> DeleteRowsAsync(int sheetId, int startRowIndex, int endRowIndex)
        {
            var requests = new List<Request>
            {
                new Request
                {
                    DeleteDimension = new DeleteDimensionRequest
                    {
                        Range = new DimensionRange
                        {
                            SheetId = sheetId,
                            Dimension = "ROWS",
                            StartIndex = startRowIndex,
                            EndIndex = endRowIndex
                        }
                    }
                }
            };

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
            var response = await _sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();
            return response.Replies != null;
        }

        /// <summary>
        /// Get all sheets in the spreadsheet
        /// </summary>
        /// <returns>A dictionary of sheet names and their IDs</returns>
        public async Task<Dictionary<string, int>> GetSheetsAsync()
        {
            var response = await _sheetsService.Spreadsheets.Get(_spreadsheetId).ExecuteAsync();
            var sheets = new Dictionary<string, int>();
            
            foreach (var sheet in response.Sheets)
            {
                if (sheet.Properties?.SheetId != null && !string.IsNullOrEmpty(sheet.Properties?.Title))
                {
                    sheets.Add(sheet.Properties.Title, sheet.Properties.SheetId.Value);
                }
            }
            
            return sheets;
        }
    }
}
