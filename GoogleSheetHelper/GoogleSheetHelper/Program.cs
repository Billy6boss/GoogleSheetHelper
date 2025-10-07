using GoogleSheetHelper;
using Microsoft.Extensions.Configuration;

// 主程式進入點
await RunGoogleSheetExampleAsync();

static async Task RunGoogleSheetExampleAsync()
{
    try
    {
        Console.WriteLine("Google Sheets API .NET 9 Example");
        Console.WriteLine("---------------------------------");

        // 從 appsettings.json 載入設定
        IConfiguration configuration = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        // 從設定檔中取得值
        string credentialsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, configuration["GoogleSheet:CredentialsPath"] ?? "credentials.json");
        string spreadsheetId = configuration["GoogleSheet:SpreadsheetId"] ?? throw new InvalidOperationException("SpreadsheetId not found in appsettings.json");

        // 檢查憑證檔是否存在
        if (File.Exists(credentialsPath) == false)
        {
            Console.WriteLine($"Error: Credentials file not found at {credentialsPath}");
            Console.WriteLine("Please follow the setup instructions to create and download your credentials file.");
            return;
        }

        // 初始化服務
        var sheetService = new GoogleSheetService(credentialsPath, spreadsheetId);

        // 獲取所有工作表並讓使用者選擇一個
        Console.WriteLine("\nGetting all sheets in the spreadsheet...");
        var sheets = await sheetService.GetSheetsAsync();

        if (sheets.Count == 0)
        {
            Console.WriteLine("No sheets found in the spreadsheet.");
            return;
        }

        // 顯示可用的工作表
        Console.WriteLine("\nAvailable sheets:");
        var sheetsList = new List<string>(sheets.Keys);
        for (int i = 0; i < sheetsList.Count; i++)
        {
            Console.WriteLine($"[{i + 1}] {sheetsList[i]}");
        }

        // 讓使用者選擇一個工作表
        string selectedSheet = GetUserSelectedSheet(sheetsList);
        Console.WriteLine($"\nYou selected: {selectedSheet}");

        bool continueOperating = true;
        while (continueOperating)
        {
            // 顯示操作選單
            Console.WriteLine("\nChoose an operation:");
            Console.WriteLine("[1] Read data");
            Console.WriteLine("[2] Append data");
            Console.WriteLine("[3] Update data");
            Console.WriteLine("[4] Clear data");
            Console.WriteLine("[5] Select a different sheet");
            Console.WriteLine("[6] Exit");

            Console.Write("\nEnter your choice (1-6): ");
            string? choice = Console.ReadLine();

            switch (choice)
            {
                case "1": // 讀取資料
                    await ReadSheetDataAsync(sheetService, selectedSheet);
                    break;

                case "2": // 附加資料
                    await AppendSheetDataAsync(sheetService, selectedSheet);
                    break;

                case "3": // 更新資料
                    await UpdateSheetDataAsync(sheetService, selectedSheet);
                    break;

                case "4": // 清除資料
                    await ClearSheetDataAsync(sheetService, selectedSheet);
                    break;

                case "5": // 選擇不同的工作表
                    selectedSheet = GetUserSelectedSheet(sheetsList);
                    Console.WriteLine($"\nYou selected: {selectedSheet}");
                    break;

                case "6": // 退出
                    continueOperating = false;
                    break;

                default:
                    Console.WriteLine("Invalid choice. Please try again.");
                    break;
            }
        }

        Console.WriteLine("\nProgram completed successfully!");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
        Console.WriteLine(ex.StackTrace);
    }
}



static string GetUserSelectedSheet(List<string> sheets)
{
    while (true)
    {
        Console.Write($"\nEnter sheet number (1-{sheets.Count}): ");
        string? input = Console.ReadLine();

        if (int.TryParse(input, out int sheetIndex) && sheetIndex >= 1 && sheetIndex <= sheets.Count)
        {
            return sheets[sheetIndex - 1];
        }

        Console.WriteLine("Invalid selection. Please try again.");
    }
}

static async Task ReadSheetDataAsync(GoogleSheetService service, string sheetName)
{
    try
    {
        Console.Write("Enter range to read (e.g., A1:D10): ");
        string? rangeInput = Console.ReadLine();

        string range = string.IsNullOrWhiteSpace(rangeInput)
            ? $"{sheetName}!A1:D5"
            : $"{sheetName}!{rangeInput}";

        Console.WriteLine($"\nReading data from '{range}'...");
        var values = await service.ReadAsync(range);

        if (values == null || values.Count == 0)
        {
            Console.WriteLine("No data found.");
            return;
        }

        Console.WriteLine("\nData:");
        for (int i = 0; i < values.Count; i++)
        {
            var row = values[i];
            Console.WriteLine($"Row {i + 1}: {string.Join(", ", row)}");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error reading data: {ex.Message}");
    }
}

static async Task AppendSheetDataAsync(GoogleSheetService service, string sheetName)
{
    try
    {
        Console.WriteLine("Enter data to append (comma-separated values, empty line to finish):");
        var rows = new List<IList<object>>();

        while (true)
        {
            string? line = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(line))
            {
                break;
            }

            var cells = line.Split(',').Select(s => (object)s.Trim()).ToList();
            rows.Add(cells);
        }

        if (rows.Count == 0)
        {
            Console.WriteLine("No data entered.");
            return;
        }

        string range = $"{sheetName}!A:Z";
        var appendedRange = await service.AppendAsync(range, rows);
        Console.WriteLine($"Data appended to range: {appendedRange}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error appending data: {ex.Message}");
    }
}

static async Task UpdateSheetDataAsync(GoogleSheetService service, string sheetName)
{
    try
    {
        Console.Write("Enter range to update (e.g., A1:B2): ");
        string? rangeInput = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(rangeInput))
        {
            Console.WriteLine("Invalid range.");
            return;
        }

        string range = $"{sheetName}!{rangeInput}";
        Console.WriteLine("Enter data to update (comma-separated values, empty line to finish):");

        var rows = new List<IList<object>>();
        while (true)
        {
            string? line = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(line))
            {
                break;
            }

            var cells = line.Split(',').Select(s => (object)s.Trim()).ToList();
            rows.Add(cells);
        }

        if (rows.Count == 0)
        {
            Console.WriteLine("No data entered.");
            return;
        }

        var updatedCells = await service.UpdateAsync(range, rows);
        Console.WriteLine($"Updated {updatedCells} cells");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error updating data: {ex.Message}");
    }
}

static async Task ClearSheetDataAsync(GoogleSheetService service, string sheetName)
{
    try
    {
        Console.Write("Enter range to clear (e.g., A1:D10): ");
        string? rangeInput = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(rangeInput))
        {
            Console.WriteLine("Invalid range.");
            return;
        }

        string range = $"{sheetName}!{rangeInput}";
        Console.WriteLine($"\nClearing data from '{range}'...");

        string clearedRange = await service.ClearAsync(range);
        Console.WriteLine($"Cleared range: {clearedRange}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error clearing data: {ex.Message}");
    }
}