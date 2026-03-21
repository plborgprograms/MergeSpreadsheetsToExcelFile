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

            }

            package.SaveAs(new FileInfo(outputFile));
        }

        Console.WriteLine("Excel workbook created at: " + outputFile);
    }
}