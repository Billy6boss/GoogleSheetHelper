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
        /// 使用 JSON 憑證檔初始化 Google Sheets 服務
        /// </summary>
        /// <param name="credentialsFilePath">JSON 憑證檔的路徑</param>
        /// <param name="spreadsheetId">試算表的 ID (可在 URL 中找到)</param>
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
        /// 從試算表中讀取特定範圍的資料
        /// </summary>
        /// <param name="range">範圍的 A1 表示法 (例如, "Sheet1!A1:D10")</param>
        /// <returns>行列表，其中每一行是一個儲存格值的列表</returns>
        public async Task<IList<IList<object>>> ReadAsync(string range)
        {
            var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = await request.ExecuteAsync();
            return response.Values ?? new List<IList<object>>();
        }

        /// <summary>
        /// 更新試算表中特定範圍的資料
        /// </summary>
        /// <param name="range">範圍的 A1 表示法 (例如, "Sheet1!A1:D10")</param>
        /// <param name="values">要更新的值</param>
        /// <returns>更新的儲存格數量</returns>
        public async Task<int> UpdateAsync(string range, IList<IList<object>> values)
        {
            var valueRange = new ValueRange { Values = values };
            var request = _sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            var response = await request.ExecuteAsync();
            return response.UpdatedCells ?? 0;
        }

        /// <summary>
        /// 向試算表中特定範圍附加資料
        /// </summary>
        /// <param name="range">範圍的 A1 表示法 (例如, "Sheet1!A:D")</param>
        /// <param name="values">要附加的值</param>
        /// <returns>更新後的範圍</returns>
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
        /// 清除試算表中特定範圍的資料
        /// </summary>
        /// <param name="range">範圍的 A1 表示法 (例如, "Sheet1!A1:D10")</param>
        /// <returns>已清除的範圍</returns>
        public async Task<string> ClearAsync(string range)
        {
            var request = _sheetsService.Spreadsheets.Values.Clear(new ClearValuesRequest(), _spreadsheetId, range);
            var response = await request.ExecuteAsync();
            return response.ClearedRange ?? string.Empty;
        }

        /// <summary>
        /// 從試算表中刪除列
        /// </summary>
        /// <param name="sheetId">工作表 ID (不是試算表 ID)</param>
        /// <param name="startRowIndex">要刪除的第一列的索引 (從 0 開始)</param>
        /// <param name="endRowIndex">要刪除的最後一列的索引 + 1 (從 0 開始)</param>
        /// <returns>成功時返回 True</returns>
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
        /// 取得試算表中的所有工作表
        /// </summary>
        /// <returns>工作表名稱和其 ID 的字典</returns>
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
