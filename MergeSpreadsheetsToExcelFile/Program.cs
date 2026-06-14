using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;

class Program
{
    static int FindColumnByHeader(ExcelWorksheet ws, int lastColumn, params string[] headerNames)
    {
        for (int col = 1; col <= lastColumn; col++)
        {
            string header = Convert.ToString(ws.Cells[1, col].Value)?.Trim() ?? string.Empty;
            foreach (string headerName in headerNames)
            {
                if (string.Equals(header, headerName, StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }
        }

        return -1;
    }

    static int AddTargetOptimizationTable(
        ExcelWorksheet ws,
        int startRow,
        string title,
        string unitLabel,
        string maxFavorableRange,
        string realizedRange)
    {
        double[] candidates = { 0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 1.75, 2.0, 2.5, 3.0, 4.0, 5.0 };

        ws.Cells[startRow, 1].Value = title;
        ws.Cells[startRow + 1, 1].Value = unitLabel;
        ws.Cells[startRow + 1, 2].Value = "Hit Rate";
        ws.Cells[startRow + 1, 3].Value = "Avg Non-Hit Realized";
        ws.Cells[startRow + 1, 4].Value = "Expected Value";

        int firstCandidateRow = startRow + 2;
        for (int i = 0; i < candidates.Length; i++)
        {
            int currentRow = firstCandidateRow + i;
            string candidateValue = candidates[i].ToString(CultureInfo.InvariantCulture);
            ws.Cells[currentRow, 1].Value = candidates[i];
            ws.Cells[currentRow, 2].Formula =
                $"=IFERROR(ROWS(FILTER({maxFavorableRange}-0, {maxFavorableRange}-0 >= {candidateValue})) / ROWS({maxFavorableRange}), 0)";
            ws.Cells[currentRow, 3].Formula =
                $"=IFERROR(AVERAGE(FILTER({realizedRange}-0, {maxFavorableRange}-0 < {candidateValue})), 0)";
            ws.Cells[currentRow, 4].Formula =
                $"=IFERROR(({ws.Cells[currentRow, 2].Address}*{candidateValue}) + ((1-{ws.Cells[currentRow, 2].Address})*{ws.Cells[currentRow, 3].Address}), 0)";
        }

        int bestRow = firstCandidateRow + candidates.Length + 1;
        int lastCandidateRow = firstCandidateRow + candidates.Length - 1;
        ws.Cells[bestRow, 1].Value = $"Best {unitLabel}:";
        ws.Cells[bestRow, 2].Formula =
            $"=IFERROR(INDEX(A{firstCandidateRow}:A{lastCandidateRow}, MATCH(MAX(D{firstCandidateRow}:D{lastCandidateRow}), D{firstCandidateRow}:D{lastCandidateRow}, 0)), \"Undefined\")";
        ws.Cells[bestRow + 1, 1].Value = $"Best {unitLabel} Expected Value:";
        ws.Cells[bestRow + 1, 2].Formula =
            $"=IFERROR(MAX(D{firstCandidateRow}:D{lastCandidateRow}), \"Undefined\")";

        return bestRow + 3;
    }

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

            if (name.EndsWith("profitResults", StringComparison.OrdinalIgnoreCase))
            {
                group = "profitResults";
            }
            else if (name.EndsWith("orderTotalsResults", StringComparison.OrdinalIgnoreCase))
            {
                group = "orderTotalsResults";
            }


            if (!groups.ContainsKey(group))
            {
                groups[group] = new List<string>();
            }
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
                            var raw = (cells[col - 1] ?? string.Empty).Trim();

                            // Try parse as number (InvariantCulture first, then CurrentCulture)
                            if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var dv) ||
                                double.TryParse(raw, NumberStyles.Any, CultureInfo.CurrentCulture, out dv))
                            {
                                ws.Cells[row, col].Value = dv;
                            }
                            else
                            {
                                ws.Cells[row, col].Value = raw;
                            }
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
                int summaryRow3 = row + 3; //profit hit profit target
                int summaryRow4 = row + 4; //profit didn't hit profit target

                // Labels in column 1
                ws.Cells[summaryRow1, 1].Value = "Avg Sell > Buy:";
                ws.Cells[summaryRow2, 1].Value = "Avg Buy > Sell:";

                // Helper to convert column number → Excel letter
                string ColLetter(int col) =>
                    OfficeOpenXml.ExcelCellAddress.GetColumnLetter(col);

                // Identify buy/sell columns
                int colBuy = lastColumn - 1;
                int colSell = lastColumn;

                //grabbing these to infer the quantity 
                string col9L = ColLetter(8); //optimistic1mRisk column
                string col10L = ColLetter(10); //totalOptimistic1mRisk column

                // Summary rows
                int avgProfitRow = row + 1;
                int avgLossRow = row + 2;
                int oddsProfitRow = row + 3;
                int oddsLossRow = row + 4;
                int reqProfitRow = row + 5;
                int targetWinRateProfitRow = row + 6;
                int targetWinRatePriceRow = row + 7;
                int targetWinRateBreakEvenRow = row + 8;

                string buyColLetter = ColLetter(colBuy);
                string sellColLetter = ColLetter(colSell);


                string buyRange = $"{buyColLetter}{dataStartRow}:{buyColLetter}{dataEndRow}";
                string sellRange = $"{sellColLetter}{dataStartRow}:{sellColLetter}{dataEndRow}";
                double commissionPerShare = 0.01;
                string commissionString = commissionPerShare.ToString(CultureInfo.InvariantCulture);
                double targetWinRate = 2.0 / 3.0;
                string targetWinRateString = targetWinRate.ToString(CultureInfo.InvariantCulture);
                string netPerShareRange = $"(({sellRange}-0)-({buyRange}-0)-{commissionString})";


                string profitTakingPricesRange = $"{ColLetter(17)}{dataStartRow}:{ColLetter(17)}{dataEndRow}"; //profitTakingPrices column

                string col9Range = $"{col9L}{dataStartRow}:{col9L}{dataEndRow}";
                string col10Range = $"{col10L}{dataStartRow}:{col10L}{dataEndRow}";


                // Labels
                ws.Cells[avgProfitRow, 1].Value = "Avg Profit (weighted):";
                ws.Cells[avgLossRow, 1].Value = "Avg Loss (weighted):";
                ws.Cells[oddsProfitRow, 1].Value = "Odds of Profit:";
                ws.Cells[oddsLossRow, 1].Value = "Odds of Loss:";
                ws.Cells[reqProfitRow, 1].Value = "Required Profit to Break Even:";
                ws.Cells[targetWinRateProfitRow, 1].Value = "Profit/Share Target for 2/3 Hit Rate:";
                ws.Cells[targetWinRatePriceRow, 1].Value = "Typical Price Target for 2/3 Hit Rate:";
                ws.Cells[targetWinRateBreakEvenRow, 1].Value = "Break-Even Profit/Share at 2/3:";

                // Weighted Avg Profit
                ws.Cells[avgProfitRow, 2].Formula =
                    $"=IFERROR(AVERAGE(FILTER(({sellRange}-0-{buyRange}-0)*(({col10Range}-0)/({col9Range}-0)), {sellRange}-0 >= {profitTakingPricesRange}-0)), \"Undefined\")";
                                 //( avg                 profit          ) *      (quantity           )    /       where profitable

                // Weighted Avg Loss
                ws.Cells[avgLossRow, 2].Formula =
                    $"=IFERROR(AVERAGE(FILTER(({buyRange}-0-{sellRange}-0)*(({col10Range}-0)/({col9Range}-0)), {profitTakingPricesRange}-0 > {sellRange}-0)), \"Undefined\")";

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
                    $"=IFERROR(ROWS(FILTER({sellRange}-0, {sellRange}-0 >= {profitTakingPricesRange}-0)) / ROWS({sellRange}-0), 0)";

                // Odds of Loss
                ws.Cells[oddsLossRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({buyRange}-0, {profitTakingPricesRange}-0 > {sellRange}-0)) / ROWS({buyRange}-0), 0)";
                // Required Profit to Break Even (based on dollars risked)
                // Formula: RequiredProfit = Risk * (1 - WinRate) / WinRate
                // Here "Risk" = average weighted loss
                ws.Cells[reqProfitRow, 2].Formula =
                    $"=IFERROR(({ws.Cells[avgLossRow, 2].Address} * (1 - {ws.Cells[oddsProfitRow, 2].Address})) / {ws.Cells[oddsProfitRow, 2].Address}, \"Undefined\")";

                // Non-circular target: this uses the net per-share move distribution,
                // not the current profitTakingLmtPrice. If this value is <= 0,
                // the data set did not support a positive 2/3 hit-rate target.
                ws.Cells[targetWinRateProfitRow, 2].Formula =
                    $"=IFERROR(PERCENTILE.INC({netPerShareRange}, 1-{targetWinRateString}), \"Undefined\")";
                ws.Cells[targetWinRatePriceRow, 2].Formula =
                    $"=IFERROR(AVERAGE({buyRange}) + {ws.Cells[targetWinRateProfitRow, 2].Address}, \"Undefined\")";
                ws.Cells[targetWinRateBreakEvenRow, 2].Formula =
                    $"=IFERROR({ws.Cells[avgLossRow, 2].Address} * (1-{targetWinRateString}) / {targetWinRateString}, \"Undefined\")";



                // Loop through all columns except the first 3
                for (int col = 4; col <= lastColumn; col++)
                {
                    string colLetter = ColLetter(col);
                    string colRange = $"{colLetter}{dataStartRow}:{colLetter}{dataEndRow}";

                    // Avg Sell > Buy (now: sell >= profit taking price for win)
                    ws.Cells[summaryRow1, col].Formula = $"=IFERROR(AVERAGE(FILTER({colRange}-0, {sellRange}-0 >= {buyRange}-0)), \"Undefined\")";

                    // Avg Buy > Sell
                    ws.Cells[summaryRow2, col].Formula = $"=IFERROR(AVERAGE(FILTER({colRange}-0, {buyRange}-0 > {sellRange}-0)), \"Undefined\")";


                    // Conditional formatting for summaryRow1
                    var rule = ws.ConditionalFormatting.AddExpression(ws.Cells[summaryRow2, col]);

                    rule.Formula = $"AND(MIN({colRange}-0)>=0, MAX({colRange}-0)<=1, ABS({colLetter}{summaryRow2}-0 - {colLetter}{summaryRow1}-0) >= 0.4)";

                    rule.Style.Font.Bold = true;


                    //profit hit profit target (sell >= profit target)
                    ws.Cells[summaryRow3, col].Formula = $"=IFERROR(AVERAGE(FILTER({colRange}-0, {sellRange}-0 >= {profitTakingPricesRange}-0)), \"Undefined\")";


                    //profit didn't hit profit target  
                    ws.Cells[summaryRow4, col].Formula = $"=IFERROR(AVERAGE(FILTER({colRange}-0, {sellRange}-0 < {profitTakingPricesRange}-0)), \"Undefined\")";

                    var ProfitTakingrule = ws.ConditionalFormatting.AddExpression(ws.Cells[summaryRow4, col]);

                    ProfitTakingrule.Formula = $"AND(MIN({colRange}-0)>=0, MAX({colRange}-0)<=1, ABS({colLetter}{summaryRow4}-0 - {colLetter}{summaryRow3}-0) >= 0.4)";

                    ProfitTakingrule.Style.Font.Bold = true;

                }


                // ---------------------------------------------
                // SUMMARY ROWS: Risk-Based Profitability Analytics
                // ---------------------------------------------

                int riskBasedSpacer = 10; //extra space before starting the risk weighted section
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
                // Risk-Based Weighted Avg Profit
                // (Sell - Buy) * (Col10 / Col9) / Risk
                // ---------------------------------------------
                ws.Cells[riskBasedavgProfitRow, 2].Formula =
                       //$"=IFERROR(AVERAGE(FILTER((({sellRange}-{buyRange})*({col10Range}/{col9Range}))/{riskPerTrade}, {sellRange}>{buyRange})), 0)";
                       $"=IFERROR(AVERAGE(FILTER(((({sellRange}-{buyRange})-0)*(({col10Range}/{col9Range})-0))/(({col10Range})-0), {sellRange}>={profitTakingPricesRange})), 0)";
                                //( avg                 profit              ) *      (quantity           )    /     (total risk) ,      where profitable

                // ---------------------------------------------
                // Risk-Based Weighted Avg Loss
                // (Buy - Sell) * (Col10 / Col9) / Risk
                // ---------------------------------------------
                ws.Cells[riskBasedavgLossRow, 2].Formula =
                    //$"=IFERROR(AVERAGE(FILTER((({buyRange}-{sellRange})*({col10Range}/{col9Range}))/{riskPerTrade}, {buyRange}>{sellRange})), 0)";
                    $"=IFERROR(AVERAGE(FILTER(((({buyRange}-{sellRange})-0)*(({col10Range}/{col9Range})-0))/(({col10Range})-0), {profitTakingPricesRange}-0 > {sellRange}-0)), 0)";


                // ---------------------------------------------
                // Odds of Profit
                // (# rows where Sell > Buy) / total rows
                // ---------------------------------------------
                ws.Cells[riskBasedoddsProfitRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({sellRange}-0, {sellRange}-0 >= {profitTakingPricesRange}-0)) / ROWS({sellRange}-0), 0)";

                // ---------------------------------------------
                // Odds of Loss
                // (# rows where Buy > Sell) / total rows
                // ---------------------------------------------
                ws.Cells[riskBasedoddsLossRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({buyRange}-0, {profitTakingPricesRange}-0 > {sellRange}-0)) / ROWS({buyRange}-0), 0)";

                // ---------------------------------------------
                // Required Profit to Break Even (Risk-Based)
                // RequiredProfit = Risk * (1 - WinRate) / WinRate
                // ---------------------------------------------
                ws.Cells[riskBasedreqProfitRow, 2].Formula =
                    $"=IFERROR(({ws.Cells[riskBasedavgLossRow, 2].Address} * (1 - {ws.Cells[oddsProfitRow, 2].Address})) / {ws.Cells[oddsProfitRow, 2].Address}, \"Undefined\")";


                // ---------------------------------------------
                // Summary row positions
                // ---------------------------------------------
                int _5mriskBasedavgProfitRow = riskBasedSpacer + riskBasedSpacer + row + 1;
                int _5mriskBasedavgLossRow = riskBasedSpacer + riskBasedSpacer + row + 2;
                int _5mriskBasedoddsProfitRow = riskBasedSpacer + riskBasedSpacer + row + 3;
                int _5mriskBasedoddsLossRow = riskBasedSpacer + riskBasedSpacer + row + 4;
                int _5mriskBasedreqProfitRow = riskBasedSpacer + riskBasedSpacer + row + 5;
                int _5mriskBasedreqNotetRow = riskBasedSpacer + riskBasedSpacer + row + 6;

                // ---------------------------------------------
                // Labels
                // ---------------------------------------------
                ws.Cells[_5mriskBasedavgProfitRow, 1].Value = "Avg Profit (5m risk-weighted):";
                ws.Cells[_5mriskBasedavgLossRow, 1].Value = "Avg Loss (5m risk-weighted):";
                ws.Cells[_5mriskBasedoddsProfitRow, 1].Value = "Odds of Profit:";
                ws.Cells[_5mriskBasedoddsLossRow, 1].Value = "Odds of Loss:";
                ws.Cells[_5mriskBasedreqProfitRow, 1].Value = "Required Multiple Of 5m risk to break even:";
                ws.Cells[_5mriskBasedreqNotetRow, 1].Value = "Note: these are the multiples of the 5m risk for each of these answers";


                string _5mRiskcol = ColLetter(9); //optimistic5mRisk column
                string _5mTotalRiskmcol = ColLetter(11); //totalOptimistic5mRisk column
                string _5mRiskcolRange = $"{_5mRiskcol}{dataStartRow}:{_5mRiskcol}{dataEndRow}";
                string _5mTotalRiskcolRange = $"{_5mTotalRiskmcol}{dataStartRow}:{_5mTotalRiskmcol}{dataEndRow}";

                // ---------------------------------------------
                // Risk-Based Weighted Avg Profit
                // (Sell - Buy) * (Col10 / Col9) / Risk
                // ---------------------------------------------
                ws.Cells[_5mriskBasedavgProfitRow, 2].Formula =
                       //$"=IFERROR(AVERAGE(FILTER((({sellRange}-{buyRange})*({col10Range}/{col9Range}))/{riskPerTrade}, {sellRange}>{buyRange})), 0)";
                       $"=IFERROR(AVERAGE(FILTER(((({sellRange}-{buyRange})-0)*(({_5mTotalRiskcolRange}/{_5mRiskcolRange})-0))/(({_5mTotalRiskcolRange})-0), {sellRange}-0 >= {profitTakingPricesRange}-0)), 0)";
                                    //( avg                 profit              ) *      (quantity           )               /     (total risk) ,      where profitable

                // ---------------------------------------------
                // Risk-Based Weighted Avg Loss
                // (Buy - Sell) * (Col10 / Col9) / Risk
                // ---------------------------------------------
                ws.Cells[_5mriskBasedavgLossRow, 2].Formula =
                    //$"=IFERROR(AVERAGE(FILTER((({buyRange}-{sellRange})*({col10Range}/{col9Range}))/{riskPerTrade}, {buyRange}>{sellRange})), 0)";
                    $"=IFERROR(AVERAGE(FILTER(((({buyRange}-{sellRange})-0)*(({_5mTotalRiskcolRange}/{_5mRiskcolRange})-0))/(({_5mTotalRiskcolRange})-0), {profitTakingPricesRange}-0 > {sellRange}-0)), 0)";


                // ---------------------------------------------
                // Odds of Profit
                // (# rows where Sell > Buy) / total rows
                // ---------------------------------------------
                ws.Cells[_5mriskBasedoddsProfitRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({sellRange}-0, {sellRange}-0 >= {profitTakingPricesRange}-0)) / ROWS({sellRange}-0), 0)";

                // ---------------------------------------------
                // Odds of Loss
                // (# rows where Buy > Sell) / total rows
                // ---------------------------------------------
                ws.Cells[_5mriskBasedoddsLossRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({buyRange}-0, {profitTakingPricesRange}-0 > {sellRange}-0)) / ROWS({buyRange}-0), 0)";

                // ---------------------------------------------
                // Required Profit to Break Even (Risk-Based)
                // RequiredProfit = Risk * (1 - WinRate) / WinRate
                // ---------------------------------------------
                ws.Cells[_5mriskBasedreqProfitRow, 2].Formula =
                    $"=IFERROR(({ws.Cells[_5mriskBasedavgLossRow, 2].Address} * (1 - {ws.Cells[oddsProfitRow, 2].Address})) / {ws.Cells[oddsProfitRow, 2].Address}, \"Undefined\")";


                //New indicator for ema% risk 
                int colEmaClimb = FindColumnByHeader(ws, lastColumn, "realizedProfitIn5mSpread", "MovementIn5mEmaSpread");
                string emaClimbColLetter = colEmaClimb > 0 ? ColLetter(colEmaClimb) : ColLetter(29);

                string emaClimbRange = $"{emaClimbColLetter}{dataStartRow}:{emaClimbColLetter}{dataEndRow}";

                // ---------------------------------------------
                // Summary row positions
                // ---------------------------------------------
                int emaPercentBasedavgProfitRow = (riskBasedSpacer*3) + row + 1 + 1;
                int emaPercentBasedavgLossRow = (riskBasedSpacer*3) + row + 2 + 1;
                int emaPercentBasedoddsProfitRow = (riskBasedSpacer*3) + row + 3 + 1;
                int emaPercentBasedoddsLossRow = (riskBasedSpacer*3) + row + 4 + 1;
                int emaPercentBasedreqProfitRow = (riskBasedSpacer*3) + row + 5 + 1;
                // ---------------------------------------------
                // Labels
                // ---------------------------------------------
                ws.Cells[emaPercentBasedavgProfitRow, 1].Value = "Avg Profit (5m ema%):";
                ws.Cells[emaPercentBasedavgLossRow, 1].Value = "Avg Loss (5m ema%):";
                ws.Cells[emaPercentBasedoddsProfitRow, 1].Value = "Odds of Profit:";
                ws.Cells[emaPercentBasedoddsLossRow, 1].Value = "Odds of Loss:";
                ws.Cells[emaPercentBasedreqProfitRow, 1].Value = "Required Profit to Break Even:";

                // ---------------------------------------------
                // RISK PER TRADE (SET THIS VALUE)
                // ---------------------------------------------


                // ---------------------------------------------
                // EMA%-based Avg Win / Avg Loss
                // Average of (rise or fall amount) divided by initial EMA spread
                // Avg Win: average of ((Sell - Buy) / EMAspread) where Sell > Buy
                // Avg Loss: average of ((Buy - Sell) / EMAspread) where Buy > Sell
                // ---------------------------------------------
                ws.Cells[emaPercentBasedavgProfitRow, 2].Formula =
                    $"=IFERROR(AVERAGE(FILTER(({emaClimbRange}-0), {sellRange}-0 >= {profitTakingPricesRange}-0)), \"Undefined\")";

                ws.Cells[emaPercentBasedavgLossRow, 2].Formula =
                    $"=IFERROR(AVERAGE(FILTER(-1*({emaClimbRange}-0), {profitTakingPricesRange}-0 > {sellRange}-0)), \"Undefined\")";


                // ---------------------------------------------
                // Odds of Profit
                // (# rows where Sell > Buy) / total rows
                // ---------------------------------------------
                ws.Cells[emaPercentBasedoddsProfitRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({sellRange}-0, {sellRange}-0 >= {profitTakingPricesRange}-0)) / ROWS({sellRange}-0), 0)";

                // ---------------------------------------------
                // Odds of Loss
                // (# rows where Buy > Sell) / total rows
                // ---------------------------------------------
                ws.Cells[emaPercentBasedoddsLossRow, 2].Formula =
                    $"=IFERROR(ROWS(FILTER({buyRange}-0, {profitTakingPricesRange}-0 > {sellRange}-0.01)) / ROWS({buyRange}-0), 0)";

                // ---------------------------------------------
                // Required Profit to Break Even (Ema%-Based)
                // RequiredProfit = Risk * (1 - WinRate) / WinRate
                // ---------------------------------------------
                ws.Cells[emaPercentBasedreqProfitRow, 2].Formula =
                    $"=IFERROR(({ws.Cells[emaPercentBasedavgLossRow, 2].Address} * (1 - {ws.Cells[oddsProfitRow, 2].Address})) / {ws.Cells[oddsProfitRow, 2].Address}, \"Undefined\")";

                //end new indicator for ema% risk

                int maxFavorableSpreadCol = FindColumnByHeader(ws, lastColumn, "maxFavorableExcursionIn5mSpread");
                int realizedSpreadCol = FindColumnByHeader(ws, lastColumn, "realizedProfitIn5mSpread");
                int maxFavorableRCol = FindColumnByHeader(ws, lastColumn, "maxFavorableExcursionInR");
                int realizedRCol = FindColumnByHeader(ws, lastColumn, "realizedProfitInR");
                int pathBasedStartRow = (riskBasedSpacer * 4) + row + 1 + 1;

                if (maxFavorableSpreadCol > 0 && realizedSpreadCol > 0)
                {
                    string maxFavorableSpreadRange = $"{ColLetter(maxFavorableSpreadCol)}{dataStartRow}:{ColLetter(maxFavorableSpreadCol)}{dataEndRow}";
                    string realizedSpreadRange = $"{ColLetter(realizedSpreadCol)}{dataStartRow}:{ColLetter(realizedSpreadCol)}{dataEndRow}";
                    pathBasedStartRow = AddTargetOptimizationTable(
                        ws,
                        pathBasedStartRow,
                        "Path-Based Target Optimization (5m EMA Spread)",
                        "Spread Multiple",
                        maxFavorableSpreadRange,
                        realizedSpreadRange);
                }

                if (maxFavorableRCol > 0 && realizedRCol > 0)
                {
                    string maxFavorableRRange = $"{ColLetter(maxFavorableRCol)}{dataStartRow}:{ColLetter(maxFavorableRCol)}{dataEndRow}";
                    string realizedRRange = $"{ColLetter(realizedRCol)}{dataStartRow}:{ColLetter(realizedRCol)}{dataEndRow}";
                    pathBasedStartRow = AddTargetOptimizationTable(
                        ws,
                        pathBasedStartRow,
                        "Path-Based Target Optimization (Per-Buy-Setup R)",
                        "R Multiple",
                        maxFavorableRRange,
                        realizedRRange);
                }





                // Force numeric formatting
                ws.Cells[dataStartRow, 4, dataEndRow, lastColumn + 1]
                    .Style.Numberformat.Format = "0.00";

                //Scatterplot:
                // Build fully qualified ranges for chart (include sheet name and wrap in single quotes
                // in case the sheet name contains spaces)
                string xRange = $"'{ws.Name}'!{_5mTotalRiskcolRange}";

                int riskWeightedProfitCol = lastColumn + 1;
                string profitColLetter = ColLetter(riskWeightedProfitCol);

                // Build helper column
                for (int r = dataStartRow; r <= dataEndRow; r++)
                {
                    ws.Cells[r, riskWeightedProfitCol].Formula =
                         $"=IFERROR((({sellColLetter}{r}-0)-({buyColLetter}{r}-0)-{commissionString}) * (({col10L}{r}-0)/({col9L}{r}-0)), 0)";
                    //ws.Cells[r, riskWeightedProfitCol].Formula =
                    //    $"=IFERROR((({sellColLetter}{r}-0)-({buyColLetter}{r}-0)) * (({col10L}{r}-0)/({col9L}{r}-0)), 0)";
                }

                // Build chart (fully qualified ranges)
                string yRange = $"'{ws.Name}'!{profitColLetter}{dataStartRow}:{profitColLetter}{dataEndRow}";

                // Create a helper column to coerce the Total 5m Risk values to numeric (chart X values)
                int xHelperCol = riskWeightedProfitCol + 1;
                string xHelperColLetter = ColLetter(xHelperCol);

                for (int r = dataStartRow; r <= dataEndRow; r++)
                {
                    // Force numeric conversion of the source Total5mRisk column into the helper column
                    ws.Cells[r, xHelperCol].Formula = $"=IFERROR(({_5mTotalRiskmcol}{r}-0),0)";
                }

                string xHelperRange = $"'{ws.Name}'!{xHelperColLetter}{dataStartRow}:{xHelperColLetter}{dataEndRow}";

                var chart = ws.Drawings.AddChart("ProfitVsK", eChartType.XYScatter);
                chart.Title.Text = "Profit vs Total 5m Risk";

                // Use ExcelRange objects and their full addresses so EPPlus binds the correct ranges
                var yCells = ws.Cells[dataStartRow, riskWeightedProfitCol, dataEndRow, riskWeightedProfitCol];
                var xCells = ws.Cells[dataStartRow, xHelperCol, dataEndRow, xHelperCol];

                chart.Series.Add(yCells.FullAddress, xCells.FullAddress);

                chart.SetPosition(emaPercentBasedreqProfitRow + 3, 0, 3, 0);
                chart.SetSize(800, 500);





                //Add a new total profit/Loss row at the end based on the second to last column;
                int TotalProfitRow = pathBasedStartRow + 1;
                ws.Cells[TotalProfitRow, 1].Value = "Total Profit:";
                int totalProfitCol = riskWeightedProfitCol;
                string totalProfitColLetter = ColLetter(totalProfitCol);
                string totalProfitRange = $"{totalProfitColLetter}{dataStartRow}:{totalProfitColLetter}{dataEndRow}"; //profitTakingPrices column

                // Sum of the total profit helper column (coerce errors if any)
                ws.Cells[TotalProfitRow, 2].Formula =
                    $"=IFERROR(SUM({totalProfitRange}), \"Undefined\")";





                int newLastColumn = ws.Dimension.End.Column;
                // Convert imported data (header row + data rows) into an Excel Table so
                // the calculated summary blocks below stay outside the table.
                // Headers are assumed to be on row 1 and data from dataStartRow..dataEndRow.
                try
                {
                    if (dataEndRow >= 1 && lastColumn >= 1)
                    {
                        // sanitize table name
                        var safeName = System.Text.RegularExpressions.Regex.Replace(sheetName ?? "Sheet", "[^A-Za-z0-9_]", "_");
                        var tableName = "tbl_" + safeName;

                        var tableRange = ws.Cells[1, 1, dataEndRow, newLastColumn];
                        var table = ws.Tables.Add(tableRange, tableName);
                        table.ShowHeader = true;
                        table.ShowFilter = true;
                        table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    }
                }
                catch
                {
                    // ignore table creation failures (will not stop workbook creation)
                }

            }

            package.SaveAs(new FileInfo(outputFile));
        }

        Console.WriteLine("Excel workbook created at: " + outputFile);
    }
}
