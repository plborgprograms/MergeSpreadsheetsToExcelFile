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

                string col9L = ColLetter(8); //optimistic1mRisk column
                string col10L = ColLetter(10); //totalOptimistic1mRisk column

                // Summary rows
                int avgProfitRow = row + 1;
                int avgLossRow = row + 2;
                int oddsProfitRow = row + 3;
                int oddsLossRow = row + 4;
                int reqProfitRow = row + 5;

                string buyColLetter = ColLetter(colBuy);
                string sellColLetter = ColLetter(colSell);

                string buyRange = $"{buyColLetter}{dataStartRow}:{buyColLetter}{dataEndRow}";
                string sellRange = $"{sellColLetter}{dataStartRow}:{sellColLetter}{dataEndRow}";

                string col9Range = $"{col9L}{dataStartRow}:{col9L}{dataEndRow}";
                string col10Range = $"{col10L}{dataStartRow}:{col10L}{dataEndRow}";


                // Labels
                ws.Cells[avgProfitRow, 1].Value = "Avg Profit (weighted):";
                ws.Cells[avgLossRow, 1].Value = "Avg Loss (weighted):";
                ws.Cells[oddsProfitRow, 1].Value = "Odds of Profit:";
                ws.Cells[oddsLossRow, 1].Value = "Odds of Loss:";
                ws.Cells[reqProfitRow, 1].Value = "Required Profit to Break Even:";

                // Weighted Avg Profit
                ws.Cells[avgProfitRow, 2].Formula =
                    $"=IFERROR(AVERAGE(FILTER(({sellRange}-0-{buyRange}-0)*(({col10Range}-0)/({col9Range}-0)), {sellRange}-0 >{buyRange}-0)), \"Undefined\")";
                                 //( avg                 profit          ) *      (quantity           )    /       where profitable

                // Weighted Avg Loss
                ws.Cells[avgLossRow, 2].Formula =
                    $"=IFERROR(AVERAGE(FILTER(({buyRange}-0-{sellRange}-0)*(({col10Range}-0)/({col9Range}-0)), {buyRange}-0 >{sellRange}-0)), \"Undefined\")";

                /*
                // Odds of Profit
                ws.Cells[oddsProfitRow, 2].Formula =
                    $"=IFERROR(COUNTIF({sellRange}, \">\" & {buyRange}) / ROWS(({sellRange}-0)), 0)";

                // Odds of Loss
                ws.Cells[oddsLossRow, 2].Formula =
                    $"=IFERROR(COUNTIF({buyRange}, \">\" & {sellRange}) / ROWS({buyRange}-0), 0)";
                */
                // Odds of Profit
                ws.Cells[oddsProfitRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({sellRange}-0, {sellRange}-0 > {buyRange}-0)) / ROWS({sellRange}-0), 0)";

                // Odds of Loss
                ws.Cells[oddsLossRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({buyRange}-0, {buyRange}-0 > {sellRange}-0)) / ROWS({buyRange}-0), 0)";
                // Required Profit to Break Even (based on dollars risked)
                // Formula: RequiredProfit = Risk * (1 - WinRate) / WinRate
                // Here "Risk" = average weighted loss
                ws.Cells[reqProfitRow, 2].Formula =
                    $"=IFERROR(({ws.Cells[avgLossRow, 2].Address} * (1 - {ws.Cells[oddsProfitRow, 2].Address})) / {ws.Cells[oddsProfitRow, 2].Address}, \"Undefined\")";



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


                // ---------------------------------------------
                // SUMMARY ROWS: Risk-Based Profitability Analytics
                // ---------------------------------------------

                int riskBasedSpacer = 6; //extra space before starting the risk weighted section
                // ---------------------------------------------
                // Summary row positions
                // ---------------------------------------------
                int riskBasedavgProfitRow = riskBasedSpacer + row + 1;
                int riskBasedavgLossRow = riskBasedSpacer + row + 2;
                int riskBasedoddsProfitRow = riskBasedSpacer + row + 3;
                int riskBasedoddsLossRow = riskBasedSpacer + row + 4;
                int riskBasedreqProfitRow = riskBasedSpacer + row + 5;

                // ---------------------------------------------
                // Labels
                // ---------------------------------------------
                ws.Cells[riskBasedavgProfitRow, 1].Value = "Avg Profit (1m risk-weighted):";
                ws.Cells[riskBasedavgLossRow, 1].Value = "Avg Loss (1m risk-weighted):";
                ws.Cells[riskBasedoddsProfitRow, 1].Value = "Odds of Profit:";
                ws.Cells[riskBasedoddsLossRow, 1].Value = "Odds of Loss:";
                ws.Cells[riskBasedreqProfitRow, 1].Value = "Required Profit to Break Even:";

                // ---------------------------------------------
                // RISK PER TRADE (SET THIS VALUE)
                // ---------------------------------------------
                double riskPerTrade = 40;   // <<< YOU SET THIS (example: $40 risk per trade)


                // ---------------------------------------------
                // Risk-Based Weighted Avg Profit
                // (Sell - Buy) * (Col10 / Col9) / Risk
                // ---------------------------------------------
                ws.Cells[riskBasedavgProfitRow, 2].Formula =
                       //$"=IFERROR(AVERAGE(FILTER((({sellRange}-{buyRange})*({col10Range}/{col9Range}))/{riskPerTrade}, {sellRange}>{buyRange})), 0)";
                       $"=IFERROR(AVERAGE(FILTER(((({sellRange}-{buyRange})-0)*(({col10Range}/{col9Range})-0))/(({col10Range})-0), {sellRange}>{buyRange})), 0)";
                                //( avg                 profit              ) *      (quantity           )    /     (total risk) ,      where profitable

                // ---------------------------------------------
                // Risk-Based Weighted Avg Loss
                // (Buy - Sell) * (Col10 / Col9) / Risk
                // ---------------------------------------------
                ws.Cells[riskBasedavgLossRow, 2].Formula =
                    //$"=IFERROR(AVERAGE(FILTER((({buyRange}-{sellRange})*({col10Range}/{col9Range}))/{riskPerTrade}, {buyRange}>{sellRange})), 0)";
                    $"=IFERROR(AVERAGE(FILTER(((({buyRange}-{sellRange})-0)*(({col10Range}/{col9Range})-0))/(({col10Range})-0), {buyRange}>{sellRange})), 0)";


                // ---------------------------------------------
                // Odds of Profit
                // (# rows where Sell > Buy) / total rows
                // ---------------------------------------------
                ws.Cells[riskBasedoddsProfitRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({sellRange}-0, {sellRange}-0 > {buyRange}-0)) / ROWS({sellRange}-0), 0)";

                // ---------------------------------------------
                // Odds of Loss
                // (# rows where Buy > Sell) / total rows
                // ---------------------------------------------
                ws.Cells[riskBasedoddsLossRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({buyRange}-0, {buyRange}-0 > {sellRange}-0)) / ROWS({buyRange}-0), 0)";

                // ---------------------------------------------
                // Required Profit to Break Even (Risk-Based)
                // RequiredProfit = Risk * (1 - WinRate) / WinRate
                // ---------------------------------------------
                ws.Cells[riskBasedreqProfitRow, 2].Formula =
                    $"=IFERROR(({riskPerTrade} * (1 - {ws.Cells[oddsProfitRow, 2].Address})) / {ws.Cells[oddsProfitRow, 2].Address}, \"Undefined\")";


                // ---------------------------------------------
                // Summary row positions
                // ---------------------------------------------
                int _5mriskBasedavgProfitRow = riskBasedSpacer + riskBasedSpacer + row + 1;
                int _5mriskBasedavgLossRow = riskBasedSpacer + riskBasedSpacer + row + 2;
                int _5mriskBasedoddsProfitRow = riskBasedSpacer + riskBasedSpacer + row + 3;
                int _5mriskBasedoddsLossRow = riskBasedSpacer + riskBasedSpacer + row + 4;
                int _5mriskBasedreqProfitRow = riskBasedSpacer + riskBasedSpacer + row + 5;

                // ---------------------------------------------
                // Labels
                // ---------------------------------------------
                ws.Cells[_5mriskBasedavgProfitRow, 1].Value = "Avg Profit (5m risk-weighted):";
                ws.Cells[_5mriskBasedavgLossRow, 1].Value = "Avg Loss (5m risk-weighted):";
                ws.Cells[_5mriskBasedoddsProfitRow, 1].Value = "Odds of Profit:";
                ws.Cells[_5mriskBasedoddsLossRow, 1].Value = "Odds of Loss:";
                ws.Cells[_5mriskBasedreqProfitRow, 1].Value = "Required Profit to Break Even:";


                string _5mRiskcol= ColLetter(9); //optimistic1mRisk column
                string _5mTotalRiskmcol = ColLetter(11); //totalOptimistic1mRisk column
                string _5mRiskcolRange = $"{_5mRiskcol}{dataStartRow}:{_5mRiskcol}{dataEndRow}";
                string _5mTotalRiskcolRange = $"{_5mTotalRiskmcol}{dataStartRow}:{_5mTotalRiskmcol}{dataEndRow}";

                // ---------------------------------------------
                // Risk-Based Weighted Avg Profit
                // (Sell - Buy) * (Col10 / Col9) / Risk
                // ---------------------------------------------
                ws.Cells[_5mriskBasedavgProfitRow, 2].Formula =
                       //$"=IFERROR(AVERAGE(FILTER((({sellRange}-{buyRange})*({col10Range}/{col9Range}))/{riskPerTrade}, {sellRange}>{buyRange})), 0)";
                       $"=IFERROR(AVERAGE(FILTER(((({sellRange}-{buyRange})-0)*(({_5mTotalRiskcolRange}/{_5mRiskcolRange})-0))/(({_5mTotalRiskcolRange})-0), {sellRange}>{buyRange})), 0)";
                //( avg                 profit              ) *      (quantity           )    /     (total risk) ,      where profitable

                // ---------------------------------------------
                // Risk-Based Weighted Avg Loss
                // (Buy - Sell) * (Col10 / Col9) / Risk
                // ---------------------------------------------
                ws.Cells[_5mriskBasedavgLossRow, 2].Formula =
                    //$"=IFERROR(AVERAGE(FILTER((({buyRange}-{sellRange})*({col10Range}/{col9Range}))/{riskPerTrade}, {buyRange}>{sellRange})), 0)";
                    $"=IFERROR(AVERAGE(FILTER(((({buyRange}-{sellRange})-0)*(({_5mTotalRiskcolRange}/{_5mRiskcolRange})-0))/(({_5mTotalRiskcolRange})-0), {buyRange}>{sellRange})), 0)";


                // ---------------------------------------------
                // Odds of Profit
                // (# rows where Sell > Buy) / total rows
                // ---------------------------------------------
                ws.Cells[_5mriskBasedoddsProfitRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({sellRange}-0, {sellRange}-0 > {buyRange}-0)) / ROWS({sellRange}-0), 0)";

                // ---------------------------------------------
                // Odds of Loss
                // (# rows where Buy > Sell) / total rows
                // ---------------------------------------------
                ws.Cells[_5mriskBasedoddsLossRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({buyRange}-0, {buyRange}-0 > {sellRange}-0)) / ROWS({buyRange}-0), 0)";

                // ---------------------------------------------
                // Required Profit to Break Even (Risk-Based)
                // RequiredProfit = Risk * (1 - WinRate) / WinRate
                // ---------------------------------------------
                ws.Cells[_5mriskBasedreqProfitRow, 2].Formula =
                    $"=IFERROR(({riskPerTrade} * (1 - {ws.Cells[oddsProfitRow, 2].Address})) / {ws.Cells[oddsProfitRow, 2].Address}, \"Undefined\")";
            }

            package.SaveAs(new FileInfo(outputFile));
        }

        Console.WriteLine("Excel workbook created at: " + outputFile);
    }
}