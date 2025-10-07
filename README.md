# GoogleSheetHelper

使用 C# 操作 Google Sheet 的記錄

## .NET 9 Google Sheets API 完整設置指南 (繁體中文)

### 1. 建立 Google Cloud 專案
1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案（或選擇現有專案）
3. 導航至「APIs & Services」>「Library」
4. 搜尋「Google Sheets API」並啟用它

### 2. 建立服務帳戶憑證
1. 在 Google Cloud Console 中，前往「APIs & Services」>「Credentials」
2. 點擊「Create Credentials」>「Service Account」
3. 填寫詳細資訊並建立服務帳戶
4. 點擊新建立的服務帳戶
5. 前往「Keys」標籤頁
6. 點擊「Add Key」>「Create new key」
7. 選擇 JSON 格式並下載金鑰檔案
8. 將此檔案重命名為 `credentials.json`

### 3. 分享您的 Google Sheet
1. 開啟您的 Google Sheet
2. 點擊「共用」按鈕
3. 添加您的服務帳戶電子郵件地址（可在服務帳戶詳情中找到）
4. 授予適當的權限（如果您想修改表格，請選擇「編輯者」權限）

### 4. 將憑證檔案放入您的專案
1. 將下載的 `credentials.json` 檔案複製到您專案的輸出目錄
   - 可以將其放入專案中並設置「複製到輸出目錄」為「如果較新則複製」
   - 或在建置後將其複製到 `bin/Debug/net9.0/` 目錄

### 5. 獲取您的試算表 ID
1. 在瀏覽器中開啟您的 Google Sheet
2. 查看 URL：`https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit`
3. 從 URL 中複製 `YOUR_SPREADSHEET_ID` 部分

### 6. 設定 appsettings.json 檔案
1. 開啟 appsettings.json 檔案
2. 將 `GoogleSheet:SpreadsheetId` 的值替換為您在上一步驟中複製的 ID
3. 若有需要，可以修改 `GoogleSheet:CredentialsPath` 的值為您的憑證檔案路徑

### 7. 執行您的應用程式
建置並執行您的應用程式，測試與 Google Sheets 的連線。

## 使用 GoogleSheetService

我們創建的 `GoogleSheetService` 類別提供以下關鍵方法：

1. `ReadAsync`：從試算表的特定範圍讀取資料
2. `UpdateAsync`：更新特定範圍的資料
3. `AppendAsync`：向試算表添加新的資料列
4. `ClearAsync`：清除特定範圍的資料
5. `DeleteRowsAsync`：從試算表中刪除整列資料
6. `GetSheetsAsync`：獲取試算表中的所有工作表及其 ID

要使用此服務，只需使用您的憑證路徑和試算表 ID 建立實例：

```csharp
var service = new GoogleSheetService("path/to/credentials.json", "YOUR_SPREADSHEET_ID");
```

## 後續步驟

完成設置後：

1. 運行範例程式碼以驗證一切正常
2. 根據您的特定需求自定義 `GoogleSheetService` 類別
3. 考慮實作錯誤處理和重試邏輯以增強穩健性
4. 如果您經常讀取相同的資料，請考慮實作緩存機制
