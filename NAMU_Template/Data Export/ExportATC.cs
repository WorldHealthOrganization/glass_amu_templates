// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using NAMU_Template.Models;
using NAMU_Template.Constants;
using AMU_Template.Constants;
using AMU_Template.Helpers;



namespace NAMU_Template.Data_Export
{
    public static class ExportATC
    {
        public static void ExportATCConsumption(
        List<AtcConsumption> atcConsData,
        List<DataAvailability> availData,
        Excel.Workbook workbook)
        {
#if DEBUG
            var sTime = DateTime.Now;
            Debug.WriteLine($"Entering ExportATCConsumption: {sTime}.");
#endif
            var headers = new[]
            {
                "Country", "Year", "Sector", "Level", "ATC_Class","AM_Class", "AWR", "mEML", "PAED",
                "ATC2", "ATC3", "ATC4", "ATC5", "ROA", "DDD", "DID"
            };

            var worksheetName = "Substance Sheet";
            ProcessAndExportATCData(atcConsData, availData, workbook, headers, worksheetName);
            PopulateAWaRePivotTableAndGraphs(workbook);
            PopulateATCPivotTableAndGraphs(workbook);
#if DEBUG
            var eTime = DateTime.Now;
            var dur = eTime - sTime;
            Debug.WriteLine($"Exiting ExportATCConsumption: {eTime}. Duration=>{dur:c}.");
#endif
        }

        private static Excel.Worksheet GetChartDataSheet(Excel.Workbook workbook)
        {
            Excel.Worksheet dataSheet = null;
            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals("Aware Chart", StringComparison.OrdinalIgnoreCase))
                {
                    dataSheet = ws;
                    break;
                }
            }
            if (dataSheet == null)
            {
                dataSheet = workbook.Worksheets.Add();
                dataSheet.Name = "Aware Chart";
            }
          
            return dataSheet;
        }

        // Call this method to update the chart data sheet from the pivot table.
        private static void UpdateChartDataSheet(Excel.Workbook workbook, Excel.PivotTable pivotTable)
        {

            //pivotTable.RefreshTable();
            Excel.Worksheet dataSheet = GetChartDataSheet(workbook);
            Excel.Range ptRange = pivotTable.TableRange1;

            // Clear previous data if needed
            dataSheet.Cells.Clear();

            //// Copy pivot table values and formats into the chart data sheet.
            //Excel.Range targetRange = dataSheet.Cells[1, 1];
            //ptRange.Copy();
            //targetRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
            //targetRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
            int targetRow = 1;
            int targetCol = 1;

            foreach (Excel.Range row in ptRange.Rows)
            {
                Excel.Range firstCell = row.Cells[1, 1]; // Get the first cell in the row
                if (firstCell.Value != null && firstCell.Value.ToString() == "Grand Total")
                {
                    continue; // Skip rows with "Grand Total" in the first column
                }

                Excel.Range targetRowRange = dataSheet.Cells[targetRow, targetCol];
                row.Copy();
                targetRowRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                targetRowRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                targetRow++; // Move to the next row in the target sheet
            }
        }

        // Register the pivot table update event so that whenever the pivot table changes,
        // the chart data sheet is updated.
        // Subscribe to the SheetPivotTableUpdate event on the workbook.
        private static void RegisterPivotTableUpdateEvent(Excel.Workbook workbook, Excel.PivotTable pivotTable)
        {
            // Cast the workbook to access events.
            Excel.WorkbookEvents_Event workbookEvents = (Excel.WorkbookEvents_Event)workbook;
            workbookEvents.SheetPivotTableUpdate += new Excel.WorkbookEvents_SheetPivotTableUpdateEventHandler((object Sh, Excel.PivotTable updatedPivotTable) =>
            {
                // Check if the updated pivot table is the one we're interested in.
                if (updatedPivotTable.Name == pivotTable.Name)
                {
                    UpdateChartDataSheet(workbook, pivotTable);
                }
            });
        }

        private static void RegisterPivotTableUpdateEventForATB(Excel.Workbook workbook, Excel.PivotTable pivotTable)
        {
            // Cast the workbook to access events.
            Excel.WorkbookEvents_Event workbookEvents = (Excel.WorkbookEvents_Event)workbook;
            workbookEvents.SheetPivotTableUpdate += new Excel.WorkbookEvents_SheetPivotTableUpdateEventHandler((object Sh, Excel.PivotTable updatedPivotTable) =>
            {
                // Check if the updated pivot table is the one we're interested in.
                if (updatedPivotTable.Name == pivotTable.Name)
                {
                    UpdateChartDataSheetForATB(workbook, pivotTable);
                }
            });
        }

        public static void AddValueBasedGraph(Excel.Workbook workbook, Excel.PivotTable pivotTable)
        {
            // Retrieve or create the ChartData worksheet
            Excel.Worksheet dataSheet = GetChartDataSheet(workbook);

            // Initially populate the chart data sheet from the pivot table.
            UpdateChartDataSheet(workbook, pivotTable);

            // Determine the actual used range after updating
            Excel.Range usedRange = dataSheet.UsedRange;
            int pivotRows = usedRange.Rows.Count;
            int pivotCols = usedRange.Columns.Count;

            // Define header rows (adjust if your pivot layout changes)
            int headerRowCategories = 2; // Expected row for category labels (e.g., "A", "W", etc.)
            int headerRowSeries = 3;     // Expected row for series header ("Total DID")
            int dataStartRow = headerRowSeries + 1; // Data starts below headers

            List<(string Category, Excel.Range Range)> seriesData = new List<(string, Excel.Range)>();

            // Iterate through all columns in the used range.
            for (int col = 1; col <= pivotCols; col++)
            {
                string category = dataSheet.Cells[headerRowCategories, col].Value?.ToString().Trim() ?? "";
                string seriesHeader = dataSheet.Cells[headerRowSeries, col].Value?.ToString().Trim() ?? "";

                if ((category == "A" || category == "W" || category == "R" || category == "N" || category == "N/A") &&
                     seriesHeader.Equals("Total DID", StringComparison.OrdinalIgnoreCase))
                {
                    // Build the series range from the data start row to the last row in this column.
                    Excel.Range colRange = dataSheet.Range[
                        dataSheet.Cells[dataStartRow, col],
                        dataSheet.Cells[pivotRows, col]
                    ];
                    seriesData.Add((category, colRange));
                }
            }

            // **Sort the seriesData list to enforce the order A → W → R → N**
            Dictionary<string, int> categoryOrder = new Dictionary<string, int>
            {
                { "A", 1 },
                { "W", 2 },
                { "R", 3 },
                { "N", 4 }
            };

            seriesData.Sort((x, y) => categoryOrder[x.Category].CompareTo(categoryOrder[y.Category]));

            // Create the chart
            Excel.ChartObjects chartObjects = (Excel.ChartObjects)dataSheet.ChartObjects();
            Excel.ChartObject chartObj = chartObjects.Add(40, 150, 400, 300);
            Excel.Chart valueChart = chartObj.Chart;
            valueChart.ChartType = Excel.XlChartType.xlColumnStacked;

            // For the x-axis, assume the first column holds the "Year" values.
            Excel.Range xValues = dataSheet.Range[
                dataSheet.Cells[dataStartRow, 1],
                dataSheet.Cells[pivotRows, 1]
            ];

            // Define colors
            Dictionary<string, int> categoryColors = new Dictionary<string, int>
            {
                { "A", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green) },      // Access: Green
                { "W", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange) },     // Watch: Light Orange
                { "R", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red) },        // Reserve: Red
                { "N", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray) }        // Not classified / Not recommended: Grey
            };

            // Add each series based on the sorted order
            Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)valueChart.SeriesCollection();
            foreach (var series in seriesData)
            {
                Excel.Series newSeries = seriesCollection.NewSeries();
                newSeries.Name = $"{series.Category} Total DID";
                newSeries.Values = series.Range;
                newSeries.XValues = xValues;

                // Apply color based on category
                if (categoryColors.ContainsKey(series.Category))
                {
                    newSeries.Format.Fill.ForeColor.RGB = categoryColors[series.Category];
                }
            }

            // Format the chart
            valueChart.HasTitle = true;
            valueChart.ChartTitle.Text = "Total DID by Year";
            valueChart.HasLegend = true;

            // Fix axis issues
            Excel.Axis xAxis = valueChart.Axes(Excel.XlAxisType.xlCategory);
            xAxis.CategoryType = Excel.XlCategoryType.xlCategoryScale;
            xAxis.AxisBetweenCategories = true;
            if (xAxis.HasTitle)
            {
                xAxis.AxisTitle.Delete();
            }
            // Check if AutoFilterMode is enabled, then turn it off
            if (dataSheet.AutoFilterMode)
            {
                dataSheet.AutoFilterMode = false;
            }
        }

        public static void AddPercentageBasedGraph(Excel.Workbook workbook, Excel.PivotTable pivotTable)
        {
            // Retrieve or create the ChartData worksheet
            Excel.Worksheet dataSheet = GetChartDataSheet(workbook);

            // Initially populate the chart data sheet
            UpdateChartDataSheet(workbook, pivotTable);

            // Determine the actual used range after updating
            Excel.Range usedRange = dataSheet.UsedRange;
            int pivotRows = usedRange.Rows.Count;
            int pivotCols = usedRange.Columns.Count;

            // Convert data into a ListObject (table) for easier handling (if desired)
            Excel.ListObject table;
            try
            {
                table = dataSheet.ListObjects["AWRTable"];
                table.ShowAutoFilter = false;
            }
            catch
            {
                table = dataSheet.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    dataSheet.Range[$"A1:{(char)('A' + pivotCols - 1)}{pivotRows}"],
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing
                );
                table.Name = "AWRTable";
                table.ShowAutoFilter = false;
            }

            // Define header rows (adjust if necessary)
            int headerRowCategories = 2; // Row with category labels
            int headerRowSeries = 3;     // Row with series header ("AWR %")
            int dataStartRow = headerRowSeries + 1; // Data starts below headers

            List<(string Category, Excel.Range Range)> seriesData = new List<(string, Excel.Range)>();

            // Iterate through all columns in the used range
            for (int col = 1; col <= pivotCols - 1; col++)
            {
                // Get the category from the designated header row
                string category = dataSheet.Cells[headerRowCategories, col].Value?.ToString().Trim() ?? "";
                if (category == "A" || category == "W" || category == "R" || category == "N")
                {
                    // The percentage header is assumed to be in the next column (col + 1) in headerRowSeries
                    string seriesHeader = dataSheet.Cells[headerRowSeries, col + 1].Value?.ToString().Trim() ?? "";
                    if (seriesHeader.Equals("AWR %", StringComparison.OrdinalIgnoreCase))
                    {
                        // Build the series range from the data start row to the last row in column (col + 1)
                        Excel.Range colRange = dataSheet.Range[
                            dataSheet.Cells[dataStartRow, col + 1],
                            dataSheet.Cells[pivotRows, col + 1]
                        ];
                        seriesData.Add((category, colRange));
                    }
                }
            }

            // **Sort seriesData list to enforce order A → W → R → N**
            Dictionary<string, int> categoryOrder = new Dictionary<string, int>
            {
                { "A", 1 },
                { "W", 2 },
                { "R", 3 },
                { "N", 4 }
            };

            seriesData.Sort((x, y) => categoryOrder[x.Category].CompareTo(categoryOrder[y.Category]));

            // Create the chart
            Excel.ChartObjects chartObjects = (Excel.ChartObjects)dataSheet.ChartObjects();
            Excel.ChartObject chartObj = chartObjects.Add(580, 150, 400, 300);
            Excel.Chart chart = chartObj.Chart;
            chart.ChartType = Excel.XlChartType.xlColumnStacked;

            // Assume the first column holds the "Year" values for the x-axis
            Excel.Range xValues = dataSheet.Range[
                dataSheet.Cells[dataStartRow, 1],
                dataSheet.Cells[pivotRows, 1]
            ];

            // Define colors
            Dictionary<string, int> categoryColors = new Dictionary<string, int>
            {
                { "A", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green) },      // Access: Green
                { "W", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange) },     // Watch: Light Orange
                { "R", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red) },        // Reserve: Red
                { "N", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray) }        // Not classified / Not recommended: Grey
            };

            // Add each series based on the sorted order
            Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
            foreach (var series in seriesData)
            {
                Excel.Series newSeries = seriesCollection.NewSeries();
                newSeries.Name = $"{series.Category} AWR%";
                newSeries.Values = series.Range;
                newSeries.XValues = xValues;
                newSeries.HasDataLabels = true;

                // Ensure data labels display as percentages
                Excel.DataLabels dataLabels = newSeries.DataLabels();
                dataLabels.NumberFormat = "0.00%";

                // Apply color based on category
                if (categoryColors.ContainsKey(series.Category))
                {
                    newSeries.Format.Fill.ForeColor.RGB = categoryColors[series.Category];
                }
            }

            // Format the chart
            chart.HasTitle = true;
            chart.ChartTitle.Text = "Percentage of AWR per Year";
            chart.HasLegend = true;

            // Format Y-Axis for percentage values
            Excel.Axis yAxis = chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            yAxis.MinimumScale = 0;
            yAxis.MaximumScale = 1;
            yAxis.TickLabels.NumberFormat = "0%";

            // Format X-Axis
            Excel.Axis xAxis = chart.Axes(Excel.XlAxisType.xlCategory);
            xAxis.CategoryType = Excel.XlCategoryType.xlCategoryScale;
            xAxis.AxisBetweenCategories = true;
            if (xAxis.HasTitle)
            {
                xAxis.AxisTitle.Delete();
            }
            // Check if AutoFilterMode is enabled, then turn it off
            if (dataSheet.AutoFilterMode)
            {
                dataSheet.AutoFilterMode = false;
            }
        }

        public static void PopulateAWaRePivotTableAndGraphs(Excel.Workbook workbook)
        {
#if DEBUG
            var sTime = DateTime.Now;
            Debug.WriteLine($"Entering PopulateAWaRePivotTableAndGraphs: {sTime}.");
#endif
            Excel.Worksheet substanceSheet = workbook.Worksheets["Substance Sheet"];
            Excel.Range dataRange = substanceSheet.UsedRange;

            // Create a pivot cache and add a new sheet for the pivot table.
            Excel.PivotCache pivotCache = workbook.PivotCaches().Create(
                Excel.XlPivotTableSourceType.xlDatabase, dataRange);
            Excel.Worksheet pivotSheet = workbook.Sheets.Add();
            pivotSheet.Name = "Aware Data";

            // Unlock all cells on the sheet (optional, if you want others editable)
            pivotSheet.Cells.Locked = false;

            // Lock only cell B6.
            Excel.Range cellB6 = pivotSheet.Range["B6"];
            cellB6.Locked = true;

            // Create the pivot table.
            Excel.PivotTable pivotTable = pivotSheet.PivotTables().Add(
                pivotCache, pivotSheet.Range["A1"], "SubstancePivotTable");

            // Configure pivot fields:
            // - "Year" as row field.
            // - "AWR" as column field (this automatically creates columns for each AWR value).
            pivotTable.PivotFields("Year").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("AWR").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

            // Exclude the "N/A" value from the AWR pivot field.
            Excel.PivotField awrField = pivotTable.PivotFields("AWR");
            // Count visible items and check for "N/A"
            int visibleCount = 0;
            bool hasNAItem = false;
            foreach (Excel.PivotItem item in awrField.PivotItems())
            {
                if (item.Visible)
                    visibleCount++;
                if (item.Name.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                    hasNAItem = true;
            }

            // If "N/A" is the only visible item, show a message and skip AWR chart creation
            if (visibleCount == 1 && hasNAItem)
            {
                MessageBox.Show("The only AWR category is 'N/A'. The AWR chart will not be created.", "AWR Chart Skipped", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Proceed with other code execution without creating the AWR chart
            }
            else
            {
                // Hide "N/A" item if it's not the only item
                foreach (Excel.PivotItem item in awrField.PivotItems())
                {
                    if (item.Name.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                    {
                        item.Visible = false;
                        break; // Exit loop after hiding the item
                    }
                }

                // **Sort Pivot Field ("AWR") in Custom Order: A → W → R → N**
                Dictionary<string, int> categoryOrder = new Dictionary<string, int>
                {
                    { "A", 1 },
                    { "W", 2 },
                    { "R", 3 },
                    { "N", 4 }
                };

                List<Excel.PivotItem> pivotItems = new List<Excel.PivotItem>();
                foreach (Excel.PivotItem item in awrField.PivotItems())
                {
                    if (categoryOrder.ContainsKey(item.Name))
                    {
                        pivotItems.Add(item);
                    }
                }

                // Sort items in the desired order
                pivotItems.Sort((x, y) => categoryOrder[x.Name].CompareTo(categoryOrder[y.Name]));

                // Reorder the Pivot Items
                for (int i = 0; i < pivotItems.Count; i++)
                {
                    pivotItems[i].Position = i + 1;  // Setting position explicitly
                }

                // Add additional filters from the Substance Sheet.
                pivotTable.PivotFields("Sector").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                pivotTable.PivotFields("Level").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                pivotTable.PivotFields("PAED").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                pivotTable.PivotFields("ROA").Orientation = Excel.XlPivotFieldOrientation.xlPageField;

                // Add the "DID" field as a data field to show totals.
                Excel.PivotField didField = pivotTable.PivotFields("DID");
                Excel.PivotField totalDidField = pivotTable.AddDataField(didField, "Total DID", Excel.XlConsolidationFunction.xlSum);

                // Add the "DID" field again to show percentage distribution.
                Excel.PivotField percentDidField = pivotTable.AddDataField(didField, "AWR %", Excel.XlConsolidationFunction.xlSum);
                percentDidField.Calculation = Excel.XlPivotFieldCalculation.xlPercentOfRow;
                percentDidField.NumberFormat = "0.00%";

                // Add charts based on the pivot table.
                AddValueBasedGraph(workbook, pivotTable);
                AddPercentageBasedGraph(workbook, pivotTable);
            }

           

            // Remove grand totals for columns (optional, to reduce clutter).
            pivotTable.ColumnGrand = false;

            // Refresh the pivot table to ensure it updates properly.
            pivotTable.RefreshTable();

            // (Optional) Adjust the layout to Tabular Form to make header renaming easier.
            pivotTable.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);

            //// Add charts based on the pivot table.
            //AddValueBasedGraph(workbook, pivotTable);
            //AddPercentageBasedGraph(workbook, pivotTable);

            // Register the update event so that the chart data sheet refreshes when the pivot table is updated.
            RegisterPivotTableUpdateEvent(workbook, pivotTable);

            // Unlock all cells on the sheet.
            pivotSheet.Cells.Locked = false;

            // (Your existing code to create and refresh the pivot table goes here)

            float left = (float)cellB6.Left;
            float top = (float)cellB6.Top;
            float width = (float)cellB6.Width;
            float height = (float)cellB6.Height;

            // Add a rectangle shape to cover cell B6.
            Excel.Shape blocker = pivotSheet.Shapes.AddShape(
              Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
              left, top, width, height);


            // Set the shape's fill to be completely transparent.
            blocker.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            blocker.Fill.Solid();
            blocker.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            blocker.Fill.Transparency = 1.0f; // 100% transparent

            // Remove the border of the shape.
            blocker.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Ensure the shape is locked and cannot be moved.
            blocker.Locked = true;
            // Optionally, set its placement so it moves with cells.
            blocker.Placement = Excel.XlPlacement.xlMove;

            // Using named parameters (if available) for clarity:
            pivotSheet.Protect("AwareProtect-AWRFILTER",
                DrawingObjects: Type.Missing,
                Contents: Type.Missing,
                Scenarios: Type.Missing,
                UserInterfaceOnly: true,         // Allows code to modify the sheet
                AllowFormattingCells: false,
                AllowFormattingColumns: false,
                AllowFormattingRows: false,
                AllowInsertingColumns: false,
                AllowInsertingRows: false,
                AllowInsertingHyperlinks: false,
                AllowDeletingColumns: false,
                AllowDeletingRows: false,
                AllowSorting: false,
                AllowFiltering: true,            // Enables use of AutoFilters, if needed
                AllowUsingPivotTables: true);     // Allows interaction with PivotTable filters
#if DEBUG
            var eTime = DateTime.Now;
            var dur = eTime - sTime;
            Debug.WriteLine($"Exiting PopulateAWaRePivotTableAndGraphs: {eTime}. Duration=>{dur:c}.");
#endif
        }

        private static void UpdateChartDataSheetForATB(Excel.Workbook workbook, Excel.PivotTable pivotTable)
        {
            string sheetName = "ATB Chart";
            Excel.Worksheet dataSheet = GetChartDataSheetForATB(workbook, sheetName);
            Excel.Range ptRange = pivotTable.TableRange1;
            // int pivotRows = ptRange.Rows.Count;
            // int pivotCols = ptRange.Columns.Count;

            // Clear previous data if needed
            dataSheet.Cells.Clear();

            // Copy pivot table values and formats into the chart data sheet.
            //Excel.Range targetRange = dataSheet.Cells[1, 1];
            //ptRange.Copy();
            //targetRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
            //targetRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
            int targetRow = 1;
            int targetCol = 1;

            foreach (Excel.Range row in ptRange.Rows)
            {
                Excel.Range firstCell = row.Cells[1, 1]; // Get the first cell in the row
                if (firstCell.Value != null && firstCell.Value.ToString() == "Grand Total")
                {
                    continue; // Skip rows with "Grand Total" in the first column
                }

                Excel.Range targetRowRange = dataSheet.Cells[targetRow, targetCol];
                row.Copy();
                targetRowRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                targetRowRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                targetRow++; // Move to the next row in the target sheet
            }

        }
        private static Excel.Worksheet GetChartDataSheetForATB(Excel.Workbook workbook, string sheetName)
        {
            Excel.Worksheet dataSheet = null;
            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    dataSheet = ws;
                    break;
                }
            }
            if (dataSheet == null)
            {
                dataSheet = workbook.Worksheets.Add();
                dataSheet.Name = sheetName;
            }
            return dataSheet;
        }
        public static void AddATCGraphs(Excel.Workbook workbook, Excel.PivotTable pivotTable)
        {
            // Retrieve and update the ATC chart data sheet.
            Excel.Worksheet chartSheet = GetChartDataSheetForATB(workbook, "ATB Chart");
            UpdateChartDataSheetForATB(workbook, pivotTable);
            Excel.Range usedRange = chartSheet.UsedRange;
            int pivotRows = usedRange.Rows.Count;
            int pivotCols = usedRange.Columns.Count;

            // Exclude the "Grand Total" row if present in the first column.
            string lastRowValue = chartSheet.Cells[pivotRows, 1].Value?.ToString().Trim() ?? "";
            if (lastRowValue.Equals("Grand Total", StringComparison.OrdinalIgnoreCase))
            {
                pivotRows--; // Reduce the number of rows to exclude the Grand Total row.
            }
            // Define header rows and where data starts.
            int headerRowCategories = 2;
            int headerRowSeries = 4;
            int dataStartRow = headerRowSeries + 1;

            // Prepare two lists: one for Total DID (value-based) and one for ATC3 % (percentage-based).
            List<(string Category, Excel.Range Range)> seriesDataValue = new List<(string, Excel.Range)>();
            List<(string Category, Excel.Range Range)> seriesDataPercentage = new List<(string, Excel.Range)>();

            for (int col = 1; col <= pivotCols; col++)
            {
                // Get the ATCAMCLASS group name from row 2.
                string category = chartSheet.Cells[headerRowCategories, col].Value?.ToString().Trim() ?? "";
                // Get the ATC code from row 3.
                string code = chartSheet.Cells[3, col].Value?.ToString().Trim() ?? "";
                // Get the series header from row 4 (either "Total DID" or "ATC3 %").
                string seriesHeader = chartSheet.Cells[headerRowSeries, col].Value?.ToString().Trim() ?? "";

                // If this is a percentage series and the category is empty,
                // assume the category comes from the previous column.
                if (seriesHeader.Equals("ATC3 %", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(category) && col > 1)
                {
                    category = chartSheet.Cells[headerRowCategories, col - 1].Value?.ToString().Trim() ?? "";
                    code = chartSheet.Cells[3, col - 1].Value?.ToString().Trim() ?? "";
                }

                // Build the series range from the data start row to the last row.
                Excel.Range colRange = chartSheet.Range[
                    chartSheet.Cells[dataStartRow, col],
                    chartSheet.Cells[pivotRows, col]
                ];

                // For the value-based chart, check for "Total DID".
                if (seriesHeader.Equals("Total DID", StringComparison.OrdinalIgnoreCase))
                {
                    seriesDataValue.Add(($"{category} {code}", colRange));
                }
                // For the percentage-based chart, check for "ATC3 %".
                else if (seriesHeader.Equals("ATC3 %", StringComparison.OrdinalIgnoreCase))
                {
                    seriesDataPercentage.Add(($"{category} {code}", colRange));
                }
            }

            // Assume that the first column holds the X-axis values (e.g. Year).
            Excel.Range xValues = chartSheet.Range[
                chartSheet.Cells[dataStartRow, 1],
                chartSheet.Cells[pivotRows, 1]
            ];

            // Create the Value-Based Chart (Total DID).
            Excel.ChartObjects chartObjects = (Excel.ChartObjects)chartSheet.ChartObjects();
            Excel.ChartObject valueChartObj = chartObjects.Add(40, 150, 450, 450);
            Excel.Chart valueChart = valueChartObj.Chart;
            valueChart.ChartType = Excel.XlChartType.xlColumnStacked;
            Excel.SeriesCollection seriesCollectionValue = (Excel.SeriesCollection)valueChart.SeriesCollection();
            foreach (var series in seriesDataValue)
            {
                Excel.Series newSeries = seriesCollectionValue.NewSeries();
                newSeries.Name = $"{series.Category} Total DID";
                newSeries.Values = series.Range;
                newSeries.XValues = xValues;
            }
            valueChart.HasTitle = true;
            valueChart.ChartTitle.Text = "Total DID by ATCAMClass & ATC3";
            valueChart.HasLegend = true;
            valueChart.Legend.IncludeInLayout = true;
            valueChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
            valueChart.Legend.Font.Size = 10; // Adjust font size if necessary

            // Create the Percentage-Based Chart (ATC3 %).
            Excel.ChartObject percentageChartObj = chartObjects.Add(580, 150, 450, 450);
            Excel.Chart percentageChart = percentageChartObj.Chart;
            percentageChart.ChartType = Excel.XlChartType.xlColumnStacked;
            Excel.SeriesCollection seriesCollectionPercentage = (Excel.SeriesCollection)percentageChart.SeriesCollection();
            foreach (var series in seriesDataPercentage)
            {
                Excel.Series newSeries = seriesCollectionPercentage.NewSeries();
                newSeries.Name = $"{series.Category} ATC3 %";
                newSeries.Values = series.Range;
                newSeries.XValues = xValues;
                newSeries.HasDataLabels = true;
                // Format data labels to display as percentages.
                Excel.DataLabels dataLabels = newSeries.DataLabels();
                dataLabels.NumberFormat = "0.00%";
            }
            percentageChart.HasTitle = true;
            percentageChart.ChartTitle.Text = "Percentage ATC3 by ATCAMClass & ATC3";
            percentageChart.HasLegend = true;
            percentageChart.Legend.IncludeInLayout = true;
            percentageChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
            percentageChart.Legend.Font.Size = 10;

            // Format the Y-axis of the percentage chart for percentage values.
            Excel.Axis yAxis = percentageChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            yAxis.MinimumScale = 0;
            yAxis.MaximumScale = 1;
            yAxis.TickLabels.NumberFormat = "0%";

            // Format the X-axes for both charts.
            Excel.Axis xAxisValue = valueChart.Axes(Excel.XlAxisType.xlCategory);
            xAxisValue.CategoryType = Excel.XlCategoryType.xlCategoryScale;
            xAxisValue.AxisBetweenCategories = true;
            if (xAxisValue.HasTitle)
            {
                xAxisValue.AxisTitle.Delete();
            }
            Excel.Axis xAxisPercentage = percentageChart.Axes(Excel.XlAxisType.xlCategory);
            xAxisPercentage.CategoryType = Excel.XlCategoryType.xlCategoryScale;
            xAxisPercentage.AxisBetweenCategories = true;
            if (xAxisPercentage.HasTitle)
            {
                xAxisPercentage.AxisTitle.Delete();
            }
        }


        public static void PopulateATCPivotTableAndGraphs(Excel.Workbook workbook)
        {
            Excel.Worksheet substanceSheet = workbook.Worksheets["Substance Sheet"];
            Excel.Range dataRange = substanceSheet.UsedRange;

            Excel.PivotCache pivotCache = workbook.PivotCaches().Create(
                Excel.XlPivotTableSourceType.xlDatabase, dataRange);
            Excel.Worksheet pivotSheet = workbook.Sheets.Add();
            pivotSheet.Name = "ATB Data";

            Excel.PivotTable pivotTable = pivotSheet.PivotTables().Add(
                pivotCache, pivotSheet.Range["A1"], "ATCPivotTable");

            pivotTable.PivotFields("Year").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields("AM_Class").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

            // Get the ATCAMClass field
            Excel.PivotField atcAmClassField = pivotTable.PivotFields("AM_Class");
            // Check for the presence of "ATB"
            bool hasATBItem = false;
            foreach (Excel.PivotItem item in atcAmClassField.PivotItems())
            {
                if (item.Name.Equals("ATB", StringComparison.OrdinalIgnoreCase))
                {
                    hasATBItem = true;
                    break;
                }
            }

            // If "ATB" is not present, show a message and skip ATB chart creation
            if (!hasATBItem)
            {
                MessageBox.Show("No AM_Class code has 'ATB', so no ATB chart will be created.", "ATB Chart Skipped", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Proceed with other code execution without creating the ATB chart
            }
            else
            {
                // Hide items that are not "ATB"
                foreach (Excel.PivotItem item in atcAmClassField.PivotItems())
                {
                    if (!item.Name.Equals("ATB", StringComparison.OrdinalIgnoreCase))
                    {
                        item.Visible = false;
                    }
                }
                pivotTable.PivotFields("ATC3").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                Excel.PivotField didField = pivotTable.PivotFields("DID");
                Excel.PivotField totalDidField = pivotTable.AddDataField(didField, "Total DID", Excel.XlConsolidationFunction.xlSum);

                Excel.PivotField percentDidField = pivotTable.AddDataField(didField, "ATC3 %", Excel.XlConsolidationFunction.xlSum);
                percentDidField.Calculation = Excel.XlPivotFieldCalculation.xlPercentOfRow;
                percentDidField.NumberFormat = "0.00%";

                // Add additional filters from the Substance Sheet.
                pivotTable.PivotFields("Sector").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                pivotTable.PivotFields("Level").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                pivotTable.PivotFields("PAED").Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                pivotTable.PivotFields("ROA").Orientation = Excel.XlPivotFieldOrientation.xlPageField;

                pivotTable.RefreshTable();
                RegisterPivotTableUpdateEventForATB(workbook, pivotTable);
                // Add charts based on the pivot table.
                AddATCGraphs(workbook, pivotTable);
            }

            
            //AddATCGraphs(workbook, pivotTable);

            // After refreshing the pivot table, determine cell B6’s position.
            Excel.Range cellB6 = pivotSheet.Range["B6"];
            float left = (float)cellB6.Left;
            float top = (float)cellB6.Top;
            float width = (float)cellB6.Width;
            float height = (float)cellB6.Height;

            // Add a rectangle shape to cover cell B6.
            Excel.Shape blocker = pivotSheet.Shapes.AddShape(
              Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
              left, top, width, height);


            // Set the shape's fill to be completely transparent.
            blocker.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            blocker.Fill.Solid();
            blocker.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            blocker.Fill.Transparency = 1.0f; // 100% transparent

            // Remove the border of the shape.
            blocker.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Ensure the shape is locked and cannot be moved.
            blocker.Locked = true;
            // Optionally, set its placement so it moves with cells.
            blocker.Placement = Excel.XlPlacement.xlMove;

            // Using named parameters (if available) for clarity:
            pivotSheet.Protect("ATBProtect-AWRFILTER",
                DrawingObjects: Type.Missing,
                Contents: Type.Missing,
                Scenarios: Type.Missing,
                UserInterfaceOnly: true,         // Allows code to modify the sheet
                AllowFormattingCells: false,
                AllowFormattingColumns: false,
                AllowFormattingRows: false,
                AllowInsertingColumns: false,
                AllowInsertingRows: false,
                AllowInsertingHyperlinks: false,
                AllowDeletingColumns: false,
                AllowDeletingRows: false,
                AllowSorting: false,
                AllowFiltering: true,            // Enables use of AutoFilters, if needed
                AllowUsingPivotTables: true);     // Allows interaction with PivotTable filters
        }

        public static void ExportProductConsumption(
            List<ProductConsumption> productConsumptionData,
            List<Product> productData,
            List<DataAvailability> availData,
            Excel.Workbook workbook)
        {
            Debug.WriteLine($"Entering ExportProductConsumption: {DateTime.Now}.");
            var headers = new[]
            {
                "Country", "Year", "Sector", "Level", "Product Id", "Label", "ATC_Class","AM_Class", "AWR", "mEML", "PAED",
                "ATC2", "ATC3", "ATC4", "ATC5", "ROA", "Form","Product_Name","Ingredients","Product_Origin","Manufacturing_Country",
                "Market_Authorization_Holder","Generics","Year_Authorization","Year_Withdrawal", "PKG", "DDD", "DID"
            };

            var worksheetName = "Package Sheet";
            ProcessAndExportProductData(productConsumptionData, productData, availData, workbook, headers, worksheetName);
            Debug.WriteLine($"Exiting ExportProductConsumption: {DateTime.Now}.");
        }

        private static void ProcessAndExportATCData(
        List<AtcConsumption> atcConsumptionData,
        List<DataAvailability> availData,
        Excel.Workbook workbook,
        string[] headers,
        string worksheetName)
        {
#if DEBUG
            var sTime = DateTime.Now;
            Debug.WriteLine($"Entering ProcessAndExportATCData: {sTime}.");
#endif
            // Step 1: Parse and Filter Availability Data
            var availFiltered = new List<Tuple<string, int, HealthSector, string, DataAvailability>>();

            // Step 1: Parse and Filter Availability Data
            //var availFiltered = new List<(int Year, string Country, string Sector, string AtcClass, DataAvailability RowData)>();
            foreach (var rowData in availData)
            {
                string country = rowData.Country;
                HealthSector sector = rowData.Sector;
                int year = rowData.Year;
                string atcClass = rowData.ATCClass;
                
                bool avail = rowData.AvailabilityTotal || rowData.AvailabilityHospital || rowData.AvailabilityCommunity;
                if (avail)
                {
                    availFiltered.Add(Tuple.Create(country, year, sector, atcClass, rowData));
                }
            }

            List<HealthLevel> levels = new List<HealthLevel> { HealthLevel.Total, HealthLevel.Hospital, HealthLevel.Community };

           
            // Generate ATC data to export
            var consolidatedAtcConsData = new List<object[]>();
            foreach (var (country, year, sector, atcClass, rowData) in availFiltered)
            {
                foreach (var atcCons in atcConsumptionData.Where(p => p.Country == country && p.Year == year && p.Sector == sector && p.AtcClass == atcClass)) 
                {

                    string amClass = atcCons.AMClass;
                    string atcClass2 = atcCons.AtcClass;
                    string aware = atcCons.AWaRe;
                    YesNoNA meml = atcCons.MEML;
                    YesNoUnknown paed = atcCons.Paediatric;
                    string atc2 = atcCons.ATC2;
                    string atc3 = atcCons.ATC3;
                    string atc4 = atcCons.ATC4;
                    string atc5 = atcCons.ATC5;
                    string roa = atcCons.Roa;



                    foreach (var level in levels)
                    {
                        Decimal  dddValue = 0;
                        Decimal didValue = 0;
                        Decimal pkgValue = 0;
                        // Retrieve the appropriate data based on sector and level
                        if (level == HealthLevel.Total && atcCons.AvailabilityTotal)
                        {
                            dddValue = atcCons.DDDConsumptionTotal;
                            didValue = atcCons.DIDConsumptionTotal;
                            pkgValue = atcCons.PKGConsumptionTotal;
                        }
                        else
                        {
                            if (level == HealthLevel.Hospital && atcCons.AvailabilityHospital)
                            {
                                dddValue = atcCons.DDDConsumptionHospital;
                                didValue = atcCons.DIDConsumptionHospital;
                                pkgValue = atcCons.PKGConsumptionHospital;
                            }
                            else
                            {
                                if (level == HealthLevel.Community && atcCons.AvailabilityCommunity)
                                {
                                    dddValue = atcCons.DDDConsumptionCommunity;
                                    didValue = atcCons.DIDConsumptionCommunity;
                                    pkgValue = atcCons.PKGConsumptionCommunity;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }

                        // Add consolidated data row
                        var datum = new object[]
                        {
                            country,                                                                        // "Country"
                            year,                                                                           // "Year"
                            HealthSectorLevelString.GetStringForHealthSector(sector),                      // "Sector"
                            HealthSectorLevelString.GetStringForHealthLevel(level),                        // "Level"
                            atcCons.AtcClass,                                                               // "ATCAMClass"
                            atcCons.AMClass,                                                                // "AMClass"
                            atcCons.AWaRe,                                                                  // "AWR"
                            YesNoNAString.GetStringFromYesNoNA(atcCons.MEML),                                                         // "mEML"
                            YesNoUnknownString.GetStringFromYesNoUnk(atcCons.Paediatric),              // "PAED"
                            atcCons.ATC2,                                                                   // "ATC2"
                            atcCons.ATC3,                                                                   // "ATC3"
                            atcCons.ATC4,                                                                   // "ATC4"
                            atcCons.ATC5,                                                                   // "ATC5"
                            atcCons.Roa,                                                                    // "ROA"
                            dddValue,                                                                       // "DDD"
                            didValue                                                                        // "DID"
                        };
                        consolidatedAtcConsData.Add(datum);
                    }
                }
            }

            // Step 4: Export Data in Batches
            var worksheet = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(ws => ws.Name == worksheetName)
                            ?? workbook.Worksheets.Add();
            worksheet.Name = worksheetName;

            // Write headers
            for (int col = 0; col < headers.Length; col++)
                worksheet.Cells[1, col + 1].Value = headers[col];

            // Write data in batches
            int batchSize = 1000; // Number of rows to write in each batch
            int row = 2;
            for (int i = 0; i < consolidatedAtcConsData.Count; i += batchSize)
            {
                var batch = consolidatedAtcConsData.Skip(i).Take(batchSize).ToArray(); // Jagged array (object[][])
                var batch2D = DataHelper.ConvertTo2DArray(batch); // Convert to 2D array (object[,])

                var range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row + batch2D.GetLength(0) - 1, headers.Length]];
                range.Value = batch2D; // Write the 2D array to the range
                row += batch2D.GetLength(0);
            }
#if DEBUG
            var eTime = DateTime.Now;
            var dur = eTime - sTime;
            Debug.WriteLine($"Exiting ProcessAndExportATCData: {eTime}. Duration=>{dur:c}.");
#endif
        }
        private static void ProcessAndExportProductData(
            List<ProductConsumption> prodConsumptionData,
            List<Product> productList,
            List<DataAvailability> availData,
            Excel.Workbook workbook,
            string[] headers,
            string worksheetName)
        {

            Dictionary<string, Product> productDict = productList.ToDictionary(pr => pr.UniqueId);


            var availFiltered = new List<Tuple<string, int, HealthSector, string, DataAvailability>>();

            // Step 1: Parse and Filter Availability Data
            //var availFiltered = new List<(int Year, string Country, string Sector, string AtcClass, DataAvailability RowData)>();
            foreach (var rowData in availData)
            {
                string country = rowData.Country;
                HealthSector sector = rowData.Sector;
                int year = rowData.Year;
                string atcClass = rowData.ATCClass;
                bool avail = rowData.AvailabilityTotal || rowData.AvailabilityHospital || rowData.AvailabilityCommunity;
                if (avail)
                {
                    availFiltered.Add(Tuple.Create(country, year, sector, atcClass, rowData));
                }
            }

            // Step 2: Parse and Map Product Data
            var prodConsList = prodConsumptionData.Where(p => !string.IsNullOrEmpty(p.ProductId)).ToList();

            // Step 3: Consolidate Data
            var consolidatedData = new List<object[]>(); // Store rows as arrays for batch writing
            var uniqueKeys = new HashSet<(string Country, int Year, string ProductUniqueId, HealthSector Sector, HealthLevel Level)>();

            List<HealthLevel> levels = new List<HealthLevel> { HealthLevel.Total, HealthLevel.Hospital, HealthLevel.Community };

            foreach (var (country, year, sector, atcClass, cysaAvailData) in availFiltered)
            {
                // Match product data for the current sector and availability
                foreach (var prodCons in prodConsList.Where(p => p.Country==country && p.Year==year && p.Sector == sector && p.AtcClass== atcClass)) 
                {
                    string prodId = prodCons.ProductId;
                    string label = prodCons.Label;
                    string amClass = prodCons.AMClass;
                    string atcClass2 = prodCons.AtcClass;
                    string aware = prodCons.AWaRe;
                    YesNoNA meml = prodCons.MEML;
                    YesNoUnknown paed = prodCons.Paediatric;
                    string atc2 = prodCons.ATC2;
                    string atc3 = prodCons.ATC3;
                    string atc4 = prodCons.ATC4;
                    string atc5 = prodCons.ATC5;
                    string roa = prodCons.Roa;

                    Product pr = productDict[prodCons.ProductUniqueId];

                    foreach (var level in levels)
                    {
                        decimal dddValue = 0;
                        decimal didValue = 0;
                        decimal pkgValue = 0;
                        // Retrieve the appropriate data based on sector and level
                        if (level == HealthLevel.Total && prodCons.AvailabilityTotal)
                        {
                            dddValue = prodCons.DDDConsumptionTotal;
                            didValue = prodCons.DIDConsumptionTotal;
                            pkgValue = prodCons.PKGConsumptionTotal;
                        }
                        else
                        {
                            if (level == HealthLevel.Hospital && prodCons.AvailabilityHospital)
                            {
                                dddValue = prodCons.DDDConsumptionHospital;
                                didValue = prodCons.DIDConsumptionHospital;
                                pkgValue = prodCons.PKGConsumptionHospital;
                            }
                            else
                            {
                                if (level == HealthLevel.Community && prodCons.AvailabilityCommunity)
                                {
                                    dddValue = prodCons.DDDConsumptionCommunity;
                                    didValue = prodCons.DIDConsumptionCommunity;
                                    pkgValue = prodCons.PKGConsumptionCommunity;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }

                                              


                        // Add consolidated data row
                        //var datum = new object[]
                        //{
                        //    country,                                                                        // "Country"
                        //    year,                                                                           // "Year"
                        //    HealthSectorLevelString.GetStringForHealthSector(sector),                       // "Sector"
                        //    HealthSectorLevelString.GetStringForHealthLevel(level),                         // "Level"
                        //    prodCons.ProductId,                                                             // "Product Id"
                        //    prodCons.Label,                                                                 // "Label"
                        //    atcClass2,                                                                      // "AMClass"
                        //    amClass,                                                                        // "ATCAMClass"
                        //    prodCons.AWaRe,                                                                 // "AWR"
                        //    prodCons.MEML?"YES":"NO",                                                          // "mEML"
                        //    YesNoUnknownString.getStringFromYNUFromString(prodCons.Paediatric),             // "PAED"
                        //    prodCons.ATC2,                                                                  // "ATC2"
                        //    prodCons.ATC3,                                                                  // "ATC3"
                        //    prodCons.ATC4,                                                                  // "ATC4"
                        //    atc5,                                                                           // "ATC5"
                        //    prodCons.Roa,                                                                   // "ROA"
                        //    pkgValue,                                                                       // "PKG"
                        //    dddValue,                                                                       // "DDD"
                        //    didValue                                                                        // "DID"
                        //};
                        //consolidatedData.Add(datum);

                        consolidatedData.Add(CreateProductConsumptionArrayData(
                            country, 
                            year, 
                            HealthSectorLevelString.GetStringForHealthSector(sector), 
                            HealthSectorLevelString.GetStringForHealthLevel(level),
                            prodCons.ProductId,
                            prodCons.Label,
                            atcClass2,
                            amClass,
                            prodCons.AWaRe,
                            YesNoNAString.GetStringFromYesNoNA(prodCons.MEML),
                            YesNoUnknownString.GetStringFromYesNoUnk(prodCons.Paediatric),
                            prodCons.ATC2,
                            prodCons.ATC3,
                            prodCons.ATC4,
                            atc5,
                            prodCons.Roa,
                            pkgValue,
                            dddValue,
                            didValue,
                            pr
                         ));

                    }               
                }
            }

            // Step 4: Export Data to Excel
            Excel.Worksheet worksheet = null;

            // Try to get the worksheet by name
            try
            {
                worksheet = workbook.Worksheets[worksheetName] as Excel.Worksheet;
            }
            catch
            {
                // If an exception occurs (worksheet not found), create a new one
                worksheet = workbook.Worksheets.Add() as Excel.Worksheet;
                worksheet.Name = worksheetName;  // Set the name of the new worksheet
            }

            if (worksheet == null)
            {
                // Fallback, in case the worksheet still could not be accessed/created
                worksheet = workbook.Worksheets.Add() as Excel.Worksheet;
                worksheet.Name = worksheetName;  // Set the name of the new worksheet
            }

            // Write headers
            for (int col = 0; col < headers.Length; col++)
                worksheet.Cells[1, col + 1].Value = headers[col];

            // Write data in batches
            int batchSize = 1000; // Number of rows to write in each batch
            int row = 2;
            for (int i = 0; i < consolidatedData.Count; i += batchSize)
            {
                var batch = consolidatedData.Skip(i).Take(batchSize).ToArray(); // Jagged array (object[][])
                var batch2D = DataHelper.ConvertTo2DArray(batch); // Convert to 2D array (object[,])

                var range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row + batch2D.GetLength(0) - 1, headers.Length]];
                range.Value = batch2D; // Write the 2D array to the range
                row += batch2D.GetLength(0);

            }
        }

        private static object[] CreateProductConsumptionArrayData(
            string country,
            int year,
            string hs,
            string hl,
            string prodId,
            string prodLabel,
            string atcClass,
            string amClass,
            string awr,
            string meml,
            string paed,
            string atc2,
            string atc3,
            string atc4,
            string atc5,
            string roa,
            Decimal pkgs,
            Decimal ddds,
            Decimal dids,
            Product pr
            )
        {
            object[] data = new object[28]; //19 mandatory + 9 optional
            data[0] = country;
            data[1] = year;
            data[2] = hs;
            data[3] = hl;
            data[4] = prodId;
            data[5] = prodLabel;
            data[6] = atcClass;
            data[7] = amClass;
            data[8] = awr;
            data[9] = meml;
            data[10] = paed;
            data[11] = atc2;
            data[12] = atc3;
            data[13] = atc4;
            data[14] = atc5;
            data[15] = roa;
            data[25] = pkgs;
            data[26] = ddds;
            data[27] = dids;

            if (pr.IsProductValid())
            {
                data[16] = pr.GetValueForVariable(Product.FORM_FIELD);
                data[17] = pr.GetValueForVariable(Product.PRODUCT_NAME_FIELD);
                data[18] = pr.GetValueForVariable(Product.INGREDIENTS_FIELD);
                data[19] = pr.GetValueForVariable(Product.PRODUCT_ORIGIN_FIELD);
                data[20] = pr.GetValueForVariable(Product.MANUFACTURER_COUNTRY_FIELD);
                data[21] = pr.GetValueForVariable(Product.MARKET_AUTH_HOLDER_FIELD);
                data[22] = pr.GetValueForVariable(Product.GENERICS_FIELD);
                object y = pr.GetValueForVariable(Product.YEAR_AUTHORIZATION_FIELD);
                if ((int)y != 0)
                {
                    data[23] = y;
                }
                y = pr.GetValueForVariable(Product.YEAR_WITHDRAWAL_FIELD);
                if ((int)y != 0)
                {
                    data[24] = y;
                }
            }

            return data;
        }

        //private static object[,] ConvertTo2DArray(object[][] jaggedArray)
        //{
        //    int rows = jaggedArray.Length;
        //    int cols = jaggedArray[0].Length;
        //    object[,] result = new object[rows, cols];

        //    for (int i = 0; i < rows; i++)
        //    {
        //        for (int j = 0; j < cols; j++)
        //        {
        //            result[i, j] = jaggedArray[i][j];
        //        }
        //    }

        //    return result;
        //}
    }
}
