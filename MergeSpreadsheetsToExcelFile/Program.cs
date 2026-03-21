using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // Set EPPlus license context (required)
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        string inputDir =  @"C:\Eclipse-workspace\TWS API\samples\Cpp\Output Data";
        string outputDir = @"C:\Eclipse-workspace\TWS API\samples\Cpp\MergedCsvs";
        string outputFile = Path.Combine(outputDir, "MergedOutput.xlsx");

        var csvFiles = Directory.GetFiles(inputDir, "*.csv");

        // Group files by suffix after last underscore
        var groups = new Dictionary<string, List<string>>();

        Console.WriteLine("Grouping Files");
        foreach (var csv in csvFiles)
        {
            string name = Path.GetFileNameWithoutExtension(csv);
            string group = name.Split('_').Last();   // e.g., "Entries"

            if (!groups.ContainsKey(group))
                groups[group] = new List<string>();

            groups[group].Add(csv);
        }



        Console.WriteLine("Creating Excel workbook created at: " + outputFile);

        using (var package = new ExcelPackage())
        {
            foreach (var group in groups)
            {
                string sheetName = group.Key;
                Console.WriteLine("Creating Sheet: " + sheetName);
                // Excel sheet name limit
                if (sheetName.Length > 31)
                    sheetName = sheetName.Substring(sheetName.Length - 31);

                var ws = package.Workbook.Worksheets.Add(sheetName);

                int row = 1;
                bool firstFile = true;

                foreach (var csvPath in group.Value)
                {
                    Console.WriteLine("Importing csvPath: " + csvPath);

                    int lineIndex = 0;  // line number within this CSV

                    foreach (var line in File.ReadLines(csvPath))
                    {
                        lineIndex++;

                        // Skip headers: first file vs subsequent files
                        if (firstFile && lineIndex <= 2)   // skip first 2 lines of first file
                            continue;

                        if (!firstFile && lineIndex <= 3)  // skip first 3 lines of subsequent files
                            continue;

                        var cells = line.Split(',');

                        for (int col = 1; col <= cells.Length; col++)
                        {
                            ws.Cells[row, col].Value = cells[col - 1];
                        }

                        row++;
                    }

                    firstFile = false;
                }

                /*
                // ---------------------------------------------
                // ADD SUMMARY ROWS (Avg Sell > Buy, Avg Buy > Sell)
                // ---------------------------------------------

                // Data ends at row-1
                int dataStartRow = 1;   // or 2 if you have a header row in Excel
                int dataEndRow = row - 1;

                // Determine last two columns dynamically
                int lastColumn = ws.Dimension.End.Column;
                int colBuy = lastColumn - 1;
                int colSell = lastColumn;

                // Convert column numbers to Excel letters
                string ColLetter(int col) =>
                    OfficeOpenXml.ExcelCellAddress.GetColumnLetter(col);

                string buyColLetter = ColLetter(colBuy);
                string sellColLetter = ColLetter(colSell);

                // Build ranges like F1:F5000
                string buyRange = $"{buyColLetter}{dataStartRow}:{buyColLetter}{dataEndRow}";
                string sellRange = $"{sellColLetter}{dataStartRow}:{sellColLetter}{dataEndRow}";

                // Summary rows
                int summaryRow1 = row + 1;
                int summaryRow2 = row + 2;

                // Labels
                ws.Cells[summaryRow1, 1].Value = "Avg Sell > Buy:";
                ws.Cells[summaryRow2, 1].Value = "Avg Buy > Sell:";

                // Formulas
                ws.Cells[summaryRow1, 2].Formula = $"=AVERAGE(FILTER({sellRange}, {sellRange} > {buyRange}))";

                ws.Cells[summaryRow2, 2].Formula = $"=AVERAGE(FILTER({buyRange}, {buyRange} > {sellRange}))";
                */

                // ---------------------------------------------
                // ADD SUMMARY ROWS FOR ALL COLUMNS EXCEPT FIRST 3
                // ---------------------------------------------

                int dataStartRow = 2;          // or 2 if you have headers in Excel
                int dataEndRow = row - 1;

                int lastColumn = ws.Dimension.End.Column;

                // Summary rows
                int summaryRow1 = row + 1;
                int summaryRow2 = row + 2;

                // Labels in column 1
                ws.Cells[summaryRow1, 1].Value = "Avg Sell > Buy:";
                ws.Cells[summaryRow2, 1].Value = "Avg Buy > Sell:";

                // Helper to convert column number → Excel letter
                string ColLetter(int col) =>
                    OfficeOpenXml.ExcelCellAddress.GetColumnLetter(col);

                // Identify buy/sell columns
                int colBuy = lastColumn - 1;
                int colSell = lastColumn;

                string buyColLetter = ColLetter(colBuy);
                string sellColLetter = ColLetter(colSell);

                string buyRange = $"{buyColLetter}{dataStartRow}:{buyColLetter}{dataEndRow}";
                string sellRange = $"{sellColLetter}{dataStartRow}:{sellColLetter}{dataEndRow}";

                // Loop through all columns except the first 3
                for (int col = 4; col <= lastColumn; col++)
                {
                    string colLetter = ColLetter(col);
                    string colRange = $"{colLetter}{dataStartRow}:{colLetter}{dataEndRow}";

                    // Avg Sell > Buy (use SELL column as filter condition)
                    ws.Cells[summaryRow1, col].Formula =
                        $"=IFERROR(AVERAGE(FILTER({colRange}-0, {sellRange}-0 > {buyRange}-0)), \"Undefined\")";

                    // Avg Buy > Sell (use BUY column as filter condition)
                    ws.Cells[summaryRow2, col].Formula =
                        $"=IFERROR(AVERAGE(FILTER({colRange}-0, {buyRange}-0 > {sellRange}-0)), \"Undefined\")";
                }
            }

            package.SaveAs(new FileInfo(outputFile));
        }

        Console.WriteLine("Excel workbook created at: " + outputFile);
    }
}