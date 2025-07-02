using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using GoogleSheetHelper;

// Main program entry point
await RunGoogleSheetExampleAsync();

static async Task RunGoogleSheetExampleAsync()
{
    try
    {
        Console.WriteLine("Google Sheets API .NET 9 Example");
        Console.WriteLine("---------------------------------");

        // Replace these with your actual values
        string credentialsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "credentials.json");
        string spreadsheetId = "1EgkvAs5u7sdhe3K-ZTH55CS_wGAqFY-688nJ_nvMwio"; // Found in the URL of your Google Sheet

        // Check if credentials file exists
        if (!File.Exists(credentialsPath))
        {
            Console.WriteLine($"Error: Credentials file not found at {credentialsPath}");
            Console.WriteLine("Please follow the setup instructions to create and download your credentials file.");
            return;
        }

        // Initialize the service
        var sheetService = new GoogleSheetService(credentialsPath, spreadsheetId);

        // Demo: Read data from sheet
        Console.WriteLine("\nReading data from 'Sheet1!A1:D5'...");
        var values = await sheetService.ReadAsync("Sheet1!A1:D5");
        DisplayData(values);

        // Demo: Append data to sheet
        Console.WriteLine("\nAppending data to 'Sheet1'...");
        var newData = new List<IList<object>>
        {
            new List<object> { "Added", DateTime.Now.ToString("yyyy-MM-dd"), "via", ".NET 9" }
        };
        var appendedRange = await sheetService.AppendAsync("Sheet1!A:D", newData);
        Console.WriteLine($"Data appended to range: {appendedRange}");

        // Demo: Update data in sheet
        Console.WriteLine("\nUpdating cell 'Sheet1!A2'...");
        var updateData = new List<IList<object>>
        {
            new List<object> { "Updated" }
        };
        var updatedCells = await sheetService.UpdateAsync("Sheet1!A2", updateData);
        Console.WriteLine($"Updated {updatedCells} cells");

        // Demo: Get all sheets
        Console.WriteLine("\nGetting all sheets...");
        var sheets = await sheetService.GetSheetsAsync();
        foreach (var sheet in sheets)
        {
            Console.WriteLine($"Sheet: {sheet.Key}, ID: {sheet.Value}");
        }

        Console.WriteLine("\nExample completed successfully!");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
        Console.WriteLine(ex.StackTrace);
    }
}

static void DisplayData(IList<IList<object>> values)
{
    if (values == null || values.Count == 0)
    {
        Console.WriteLine("No data found.");
        return;
    }

    foreach (var row in values)
    {
        Console.WriteLine(string.Join(", ", row));
    }
}