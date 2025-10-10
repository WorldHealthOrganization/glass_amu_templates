using AMU_Template.Constants;
using AMU_Template.Helpers;
using HAMU_Template.Constants;
using HAMU_Template.Models;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace HAMU_Template.Data_Export
{
    public class ExporterConsumption
    {

        static string FIELD_COUNTRY = "COUNTRY";
        static string FIELD_YEAR = "YEAR";
        static string FIELD_HOSPITAL = "HOSPITAL";
        static string FIELD_LEVEL = "LEVEL";
        static string FIELD_AM_CLASS = "AM_CLASS";
        static string FIELD_ATC_CLASS = "ATC_CLASS";
        static string FIELD_AWARE = "AWARE";
        static string FIELD_MEML = "MEML";
        static string FIELD_PAEDIATRICS = "PAEDIATRICS";
        static string FIELD_ATC2 = "ATC2";
        static string FIELD_ATC3 = "ATC3";
        static string FIELD_ATC4 = "ATC4";
        static string FIELD_ATC5 = "ATC5";
        static string FIELD_ROA = "ROA";
        static string FIELD_DDD = "DDD";
        static string FIELD_DAD = "DAD";
        static string FIELD_DBD = "DBD";

        static string FIELD_TOTAL_DBD = "Total DBD";
        static string FIELD_TOTAL_DAD = "Total DAD";
        static string FIELD_TOTAL_DDD = "Total DDD";


        static string[] DataHeader = { FIELD_COUNTRY, FIELD_YEAR, FIELD_HOSPITAL, FIELD_LEVEL, FIELD_AM_CLASS, FIELD_ATC_CLASS, FIELD_AWARE, 
            FIELD_MEML, FIELD_PAEDIATRICS, FIELD_ATC2, FIELD_ATC3, FIELD_ATC4, FIELD_ATC5, FIELD_ROA, FIELD_DDD, FIELD_DAD, FIELD_DBD};


        private Excel.Worksheet Atc4Sheet;
        private Excel.PivotTable Atc4PivotTable;
        private Excel.Chart Atc4AbsoluteChart;
        private Excel.Chart Atc4RelativeChart;
        private Excel.Worksheet AWRSheet;
        private Excel.PivotTable AWRPivotTable;
        private Excel.Chart AWRAbsoluteChart;
        private Excel.Chart AWRRelativeChart;
        private Excel.Worksheet DataSheet;

        private static readonly string DATA_SHEET_NAME = "Use Data";
        private static readonly int DATA_SHEET_IDX = 1;
        private static readonly string ATC4_SHEET_NAME = "ATC4 results";
        private static readonly int ATC4_SHEET_IDX = 2;
        private static readonly string AWR_SHEET_NAME = "AWaRe results";
        private static readonly int AWR_SHEET_IDX = 3;

        private static readonly string ATC4_PIVOT_TABLE_NAME = "ATC4PivotTable";
        private static readonly string AWR_PIVOT_TABLE_NAME = "AWRPivotTable";

        private static readonly string PIVOT_TABLE_TOP_CELL = "F1";

        private static readonly string ATC4_ABSOLUTE_CHART_NAME = "ATC4 Absolute Chart";
        private static readonly string ATC4_RELATIVE_CHART_NAME = "ATC4 Relative Chart";
        private static readonly string AWR_ABSOLUTE_CHART_NAME = "AWR Absolute Chart";
        private static readonly string AWR_RELATIVE_CHART_NAME = "AWR Relative Chart";

        private static readonly string INDIC_DDD_TEXT = "DDD";
        private static readonly string INDIC_DBD_TEXT = "DDD/100 patient-days";
        private static readonly string INDIC_DAD_TEXT = "DDD/100 admissions";

        private static readonly string INDIC_DDD = "DDD";
        private static readonly string INDIC_DBD = "DBD";
        private static readonly string INDIC_DAD = "DAD";

        private static readonly List<String> INDIC_TXT_LIST = new List<string>() { INDIC_DBD_TEXT, INDIC_DAD_TEXT, INDIC_DDD_TEXT };
        private static readonly Dictionary<String, String> INDIC_TXT_INDIC_MAP = new Dictionary<string, string>()
        {
            { INDIC_DBD_TEXT,INDIC_DBD },
            { INDIC_DAD_TEXT,INDIC_DAD },
            { INDIC_DDD_TEXT,INDIC_DDD },
        };

        private static readonly Dictionary<String, String> INDIC_INDIC_TXT_MAP = new Dictionary<string, string>()
        {
            { INDIC_DBD,INDIC_DBD_TEXT },
            { INDIC_DAD, INDIC_DAD_TEXT },
            { INDIC_DDD, INDIC_DDD_TEXT },
        };

        private static readonly Dictionary<String, String> INDIC_FIELD_MAP = new Dictionary<string, string>()
        {
            { INDIC_DBD, FIELD_DBD},
            { INDIC_DAD, FIELD_DAD },
            { INDIC_DDD , FIELD_DDD },
        };

        private static readonly Dictionary<String, String> INDIC_FIELD_TOTAL_MAP = new Dictionary<string, string>()
        {
            { INDIC_DBD, FIELD_TOTAL_DBD},
            { INDIC_DAD, FIELD_TOTAL_DAD },
            { INDIC_DDD , FIELD_TOTAL_DDD },
        };

        private string SelectedATC4PivotTableIndicator;
        private string Atc4PivotCurrentIndicName;
        private string Atc4PivotCurrentTotalIndicName;
        private string Atc4PivotOldIndicName;
        private string Atc4PivotOldTotalIndicName;

        private string SelectedAWRPivotTableIndicator;
        private string AWRPivotCurrentIndicName;
        private string AWRPivotCurrentTotalIndicName;
        private string AWRPivotOldIndicName;
        private string AWRPivotOldTotalIndicName;

        public ExporterConsumption()
        {
        }

        public void ExportAtcConsumption(List<AtcConsumption> atcConsData)
        {
            Excel.Application excelApp = new Excel.Application();

            // Create a new workbook
            Excel.Workbook workbook = excelApp.Workbooks.Add();


            // Create Data Sheet
            DataSheet =  CreateDataSheet(workbook, atcConsData);

            // Create ATC4 Chart Sheet
            Atc4Sheet = CreateATC4Sheet(workbook);
            Atc4PivotTable = CreateATC4PivotTable(workbook, Atc4Sheet, DataSheet);
            CreateATC4PivotCharts(workbook, Atc4Sheet, out Atc4AbsoluteChart, out Atc4RelativeChart);

            // Create AWARE Chart Sheet
            AWRSheet = CreateAWRSheet(workbook); 
            AWRPivotTable = CreateAwrPivotTable(workbook, AWRSheet, DataSheet);
            CreateAwrPivotCharts(workbook, AWRSheet, out AWRAbsoluteChart, out AWRRelativeChart);

            excelApp.Visible = true;
            workbook.Activate();
        }

        private Excel.Worksheet CreateDataSheet(Excel.Workbook workbook, List<AtcConsumption> atcConsumptionData) {

            Excel.Worksheet ws = (Excel.Worksheet)workbook.Sheets.Add(Type.Missing, workbook.Sheets[DATA_SHEET_IDX]);
            ws.Name = DATA_SHEET_NAME;

            int startRow = 2;
            int row = startRow;

            var consolidatedData = new List<object[]>();

            var headerRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, DataHeader.Length]];
            headerRange.Value = DataHeader;

            foreach (var atcCons in atcConsumptionData)
            {
                object[] datum =
                {
                    atcCons.Country,
                    atcCons.Year,
                    atcCons.Hospital,
                    FacilityStructureLevelString.GetStringForFacilityStructureLevel(atcCons.Level),
                    atcCons.AMClass,
                    atcCons.AtcClass,
                    atcCons.AWaRe,
                    YesNoNAString.GetStringFromYesNoNA(atcCons.MEML),
                    YesNoUnknownString.GetStringFromYesNoUnk(atcCons.Paediatric),
                    atcCons.ATC2,
                    atcCons.ATC3,
                    atcCons.ATC4,
                    atcCons.ATC5,
                    atcCons.Roa,
                    atcCons.DDDConsumption,
                    atcCons.DADConsumption,
                    atcCons.DBDConsumption
                };
                consolidatedData.Add(datum);
            }


            // Write data in batches
            int batchSize = 300; // Number of rows to write in each batch
            
            for (int i = 0; i < consolidatedData.Count; i += batchSize)
            {
                var batch = consolidatedData.Skip(i).Take(batchSize).ToArray(); // Jagged array (object[][])
                var batch2D = DataHelper.ConvertTo2DArray(batch); // Convert to 2D array (object[,])

                var range = ws.Range[ws.Cells[row, 1], ws.Cells[row + batch2D.GetLength(0) - 1, DataHeader.Length]];
                range.Value = batch2D; // Write the 2D array to the range
                row += batch2D.GetLength(0);
            }
            return ws;
        }

        private Excel.Worksheet CreateATC4Sheet(Excel.Workbook workbook)
        {
            Excel.Worksheet ws = workbook.Sheets.Add(Type.Missing, workbook.Sheets[ATC4_SHEET_IDX]) as Excel.Worksheet;
            ws.Name = ATC4_SHEET_NAME;

            var cellIndicLabel = ws.Cells[1,1] as Range;
            cellIndicLabel.Value = "Indicator:";
            cellIndicLabel.Font.Bold = true;

            var validIndics = string.Join(CultureInfo.CurrentCulture.TextInfo.ListSeparator, INDIC_TXT_LIST);

            var cellIndic = ws.Cells[1, 2] as Range;
            cellIndic.Interior.Color = System.Drawing.Color.FromArgb(255, 192, 230, 245);
            cellIndic.ColumnWidth = 20;
            cellIndic.Validation.Delete();
            cellIndic.Validation.Add(
               XlDVType.xlValidateList,
               XlDVAlertStyle.xlValidAlertInformation,
               XlFormatConditionOperator.xlBetween,
               validIndics,
               Type.Missing);

            cellIndic.Validation.IgnoreBlank = true;
            cellIndic.Validation.InCellDropdown = true;
            cellIndic.Value = INDIC_DBD_TEXT;

            ws.Change += Atc4Indic_Changed;

            return ws;
        }

        private void Atc4Indic_Changed(Range target)
        {
            if (target == null || target.Address != "$B$1")
            {
                return;
            }
            var indicTxt = target.Value2 as String;
            if (!INDIC_TXT_LIST.Contains(indicTxt)) { return; }

            var indic = INDIC_TXT_INDIC_MAP[indicTxt] as String;

            Atc4PivotOldIndicName = Atc4PivotCurrentIndicName;
            Atc4PivotOldTotalIndicName = Atc4PivotCurrentTotalIndicName;

            Atc4PivotCurrentIndicName = INDIC_FIELD_MAP[indic];
            Atc4PivotCurrentTotalIndicName = INDIC_FIELD_TOTAL_MAP[indic];

            if (indic == INDIC_DBD && SelectedATC4PivotTableIndicator!=INDIC_DBD)
            {
                SelectedATC4PivotTableIndicator = INDIC_DBD;
                UpdateAtc4TableAndCharts();
            }
            else
            {
                if (indic == INDIC_DAD && SelectedATC4PivotTableIndicator != INDIC_DAD)
                {
                    SelectedATC4PivotTableIndicator = INDIC_DAD;
                    UpdateAtc4TableAndCharts();
                }
                else
                {
                    SelectedATC4PivotTableIndicator = INDIC_DDD;
                    UpdateAtc4TableAndCharts();
                }
            }
        }

        private void UpdateAtc4TableAndCharts()
        {
            var oldTotalDataField = Atc4PivotTable.PivotFields(Atc4PivotOldTotalIndicName) as PivotField;
            oldTotalDataField.Orientation = XlPivotFieldOrientation.xlHidden;

            var indicField = Atc4PivotTable.PivotFields(Atc4PivotCurrentIndicName);
            Atc4PivotTable.AddDataField(indicField, Atc4PivotCurrentTotalIndicName, Excel.XlConsolidationFunction.xlSum);

            FormatChartTitle(Atc4AbsoluteChart, "Total Use by Year", INDIC_INDIC_TXT_MAP[SelectedATC4PivotTableIndicator]);
            FormatChartTitle(Atc4RelativeChart, "Relative Use by Year", INDIC_INDIC_TXT_MAP[SelectedATC4PivotTableIndicator]);

            MoveSheetCharts(ATC4_SHEET_NAME);
        }

        private void MoveSheetCharts(string sheetName)
        {
            if (sheetName == ATC4_SHEET_NAME)
            {

                float pivotTop = (float)((Excel.Range)Atc4PivotTable.TableRange2).Top;
                float pivotHeight = (float)((Excel.Range)Atc4PivotTable.TableRange2).Height;

                float chartTop = pivotTop + pivotHeight + 20;

                ChartObject chartObj1 = Atc4Sheet.ChartObjects(ATC4_ABSOLUTE_CHART_NAME);
                chartObj1.Top = chartTop;

                ChartObject chartObj2 = Atc4Sheet.ChartObjects(ATC4_RELATIVE_CHART_NAME);
                chartObj2.Top = chartTop + chartObj1.Height + 20;
            }
            else
            {
                float pivotTop = (float)((Excel.Range)AWRPivotTable.TableRange2).Top;
                float pivotHeight = (float)((Excel.Range)AWRPivotTable.TableRange2).Height;

                float chartTop = pivotTop + pivotHeight + 20;

                ChartObject chartObj1 = AWRSheet.ChartObjects(AWR_ABSOLUTE_CHART_NAME);
                chartObj1.Top = chartTop;

                ChartObject chartObj2 = AWRSheet.ChartObjects(AWR_RELATIVE_CHART_NAME);
                chartObj2.Top = chartTop + chartObj1.Height + 20;
            }
        }

        private Excel.PivotTable CreateATC4PivotTable(Excel.Workbook workbook, Excel.Worksheet sheet, Excel.Worksheet dataSheet)
        {
            Excel.Range dataRange = dataSheet.UsedRange;
            Excel.PivotCache pivotCache = workbook.PivotCaches().Create(
                Excel.XlPivotTableSourceType.xlDatabase, dataRange);
            

            Excel.PivotTable pivotTable = sheet.PivotTables().Add(
                pivotCache, sheet.Range[PIVOT_TABLE_TOP_CELL], ATC4_PIVOT_TABLE_NAME);


            // Add the Year in row and ATC4 in colum fields.
            pivotTable.PivotFields(FIELD_YEAR).Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields(FIELD_ATC4).Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            
            // Add the data field as sum DBD.
            var indicField = pivotTable.PivotFields(FIELD_DBD) as PivotField;
            pivotTable.AddDataField(indicField, FIELD_TOTAL_DBD, Excel.XlConsolidationFunction.xlSum);

            SelectedATC4PivotTableIndicator = INDIC_DBD;
            Atc4PivotCurrentIndicName = FIELD_DBD;
            Atc4PivotCurrentTotalIndicName = FIELD_TOTAL_DBD;


            // Add additional filters.
            pivotTable.PivotFields(FIELD_AM_CLASS).Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            pivotTable.PivotFields(FIELD_ROA).Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            pivotTable.PivotFields(FIELD_PAEDIATRICS).Orientation = Excel.XlPivotFieldOrientation.xlPageField;

            // Don't show the grand total per colum
            pivotTable.ColumnGrand = false;

            pivotTable.RefreshTable();
           
            return pivotTable;
        }

        private void FormatChartTitle(Excel.Chart chart, string topTitle, string subTitle)
        {
            chart.ChartTitle.Text = topTitle + "\r" + subTitle;
            chart.ChartTitle.Format.TextFrame2.TextRange.Characters[1, topTitle.Length].Font.Size = 18;
            chart.ChartTitle.Format.TextFrame2.TextRange.Characters[topTitle.Length + 1, subTitle.Length + 1].Font.Size = 15;
            chart.ChartTitle.Format.TextFrame2.TextRange.Characters[topTitle.Length + 1, subTitle.Length + 1].Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue;
        }

        private void CreateATC4PivotCharts(Excel.Workbook workbook, Excel.Worksheet sheet, out Excel.Chart absoluteChart, out Excel.Chart relativeChart)
        {
            float pivotTop = (float)((Excel.Range)Atc4PivotTable.TableRange2).Top;
            float pivotHeight = (float)((Excel.Range)Atc4PivotTable.TableRange2).Height;

            float chartTop = pivotTop + pivotHeight + 20;

            // Aboslute chart
            ChartObject chartObj1 = sheet.ChartObjects().Add(50, chartTop, 600, 337);
            chartObj1.Name = ATC4_ABSOLUTE_CHART_NAME;
            Excel.Chart chart1 = chartObj1.Chart;

            chart1.SetSourceData(Atc4PivotTable.TableRange1, Type.Missing);
            chart1.ChartType = Excel.XlChartType.xlBarStacked;
            chart1.ShowAllFieldButtons = false;

            // Format the chart
            chart1.HasTitle = true;
            var bt = "Total Use by Year";
            var st = INDIC_INDIC_TXT_MAP[SelectedATC4PivotTableIndicator];
            FormatChartTitle(chart1, bt, st);


            // Relative chart
            ChartObject chartObj2 = sheet.ChartObjects().Add(50, chartTop + chartObj1.Height + 20, 600, 337);
            chartObj2.Name = ATC4_RELATIVE_CHART_NAME;
            Excel.Chart chart2 = chartObj2.Chart;

            chart2.SetSourceData(Atc4PivotTable.TableRange1, Type.Missing);
            chart2.ChartType = Excel.XlChartType.xlBarStacked100;
            chart2.ShowAllFieldButtons = false;

            // Format the chart
            chart2.HasTitle = true;

            bt = "Relative Use by Year";
            st = INDIC_INDIC_TXT_MAP[SelectedATC4PivotTableIndicator];
            FormatChartTitle(chart2, bt, st);

            absoluteChart = chart1;
            relativeChart = chart2;
        }

        private Excel.Worksheet CreateAWRSheet(Excel.Workbook workbook)
        {
            Excel.Worksheet ws = workbook.Sheets.Add(Type.Missing, workbook.Sheets[AWR_SHEET_IDX]) as Excel.Worksheet;
            ws.Name = AWR_SHEET_NAME;

            var cellIndicLabel = ws.Cells[1, 1] as Range;
            cellIndicLabel.Value = "Indicator:";
            cellIndicLabel.Font.Bold = true;

            var validIndics = string.Join(CultureInfo.CurrentCulture.TextInfo.ListSeparator, INDIC_TXT_LIST);

            var cellIndic = ws.Cells[1, 2] as Range;
            cellIndic.Interior.Color = System.Drawing.Color.FromArgb(255, 192, 230, 245);
            cellIndic.ColumnWidth = 20;
            cellIndic.Validation.Delete();
            cellIndic.Validation.Add(
               XlDVType.xlValidateList,
               XlDVAlertStyle.xlValidAlertInformation,
               XlFormatConditionOperator.xlBetween,
               validIndics,
               Type.Missing);

            cellIndic.Validation.IgnoreBlank = true;
            cellIndic.Validation.InCellDropdown = true;
            cellIndic.Value = INDIC_DBD_TEXT;

            ws.Change += AWRIndic_Changed;

            return ws;
        }

        private void AWRIndic_Changed(Range target)
        {
            if (target == null || target.Address != "$B$1")
            {
                return;
            }
            var indicTxt = target.Value2 as String;
            if (!INDIC_TXT_LIST.Contains(indicTxt)) { return; }

            var indic = INDIC_TXT_INDIC_MAP[indicTxt] as String;

            AWRPivotOldIndicName = AWRPivotCurrentIndicName;
            AWRPivotOldTotalIndicName = AWRPivotCurrentTotalIndicName;

            AWRPivotCurrentIndicName = INDIC_FIELD_MAP[indic];
            AWRPivotCurrentTotalIndicName = INDIC_FIELD_TOTAL_MAP[indic];

            if (indic == INDIC_DBD && SelectedAWRPivotTableIndicator != INDIC_DBD)
            {
                SelectedAWRPivotTableIndicator = INDIC_DBD;
                UpdateAWRTableAndCharts();
            }
            else
            {
                if (indic == INDIC_DAD && SelectedAWRPivotTableIndicator != INDIC_DAD)
                {
                    SelectedAWRPivotTableIndicator = INDIC_DAD;
                    UpdateAWRTableAndCharts();
                }
                else
                {
                    SelectedAWRPivotTableIndicator = INDIC_DDD;
                    UpdateAWRTableAndCharts();
                }
            }
        }

        private void UpdateAWRTableAndCharts()
        {
            var oldTotalDataField = AWRPivotTable.PivotFields(AWRPivotOldTotalIndicName) as PivotField;
            oldTotalDataField.Orientation = XlPivotFieldOrientation.xlHidden;

            var indicField = AWRPivotTable.PivotFields(AWRPivotCurrentIndicName);
            AWRPivotTable.AddDataField(indicField, AWRPivotCurrentTotalIndicName, Excel.XlConsolidationFunction.xlSum);

            FormatChartTitle(AWRAbsoluteChart, "Total Use by Year", INDIC_INDIC_TXT_MAP[SelectedAWRPivotTableIndicator]);
            FormatChartTitle(AWRRelativeChart, "Relative Use by Year", INDIC_INDIC_TXT_MAP[SelectedAWRPivotTableIndicator]);

            MoveSheetCharts(AWR_SHEET_NAME);
        }

        private Excel.PivotTable CreateAwrPivotTable(Excel.Workbook workbook, Excel.Worksheet sheet, Excel.Worksheet dataSheet)
        {
            Excel.Range dataRange = dataSheet.UsedRange;
            Excel.PivotCache pivotCache = workbook.PivotCaches().Create(
                Excel.XlPivotTableSourceType.xlDatabase, dataRange);


            Excel.PivotTable pivotTable = sheet.PivotTables().Add(
                pivotCache, sheet.Range[PIVOT_TABLE_TOP_CELL], AWR_PIVOT_TABLE_NAME);


            // Add the Year in row and ATC4 in colum fields.
            pivotTable.PivotFields(FIELD_YEAR).Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            pivotTable.PivotFields(FIELD_AWARE).Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

            // Add the data field as sum DBD.
            var indicField = pivotTable.PivotFields(FIELD_DBD) as PivotField;
            pivotTable.AddDataField(indicField, FIELD_TOTAL_DBD, Excel.XlConsolidationFunction.xlSum);

            SelectedAWRPivotTableIndicator = INDIC_DBD;
            AWRPivotCurrentIndicName = FIELD_DBD;
            AWRPivotCurrentTotalIndicName = FIELD_TOTAL_DBD;


            // Add additional filters.
            pivotTable.PivotFields(FIELD_AM_CLASS).Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            pivotTable.PivotFields(FIELD_ROA).Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            pivotTable.PivotFields(FIELD_PAEDIATRICS).Orientation = Excel.XlPivotFieldOrientation.xlPageField;

            // Don't show the grand total per colum
            pivotTable.ColumnGrand = false;

            pivotTable.RefreshTable();

            return pivotTable;
        }

        private void CreateAwrPivotCharts(Excel.Workbook workbook, Excel.Worksheet sheet, out Excel.Chart absoluteChart, out Excel.Chart relativeChart)
        {
            float pivotTop = (float)((Excel.Range)AWRPivotTable.TableRange2).Top;
            float pivotHeight = (float)((Excel.Range)AWRPivotTable.TableRange2).Height;

            float chartTop = pivotTop + pivotHeight + 20;

            // Aboslute chart
            ChartObject chartObj1 = sheet.ChartObjects().Add(50, chartTop, 600, 337);
            chartObj1.Name = AWR_ABSOLUTE_CHART_NAME;
            Excel.Chart chart1 = chartObj1.Chart;

            chart1.SetSourceData(AWRPivotTable.TableRange1, Type.Missing);
            chart1.ChartType = Excel.XlChartType.xlBarStacked;
            chart1.ShowAllFieldButtons = false;

            // Format the chart
            chart1.HasTitle = true;
            var bt = "Total Use by Year";
            var st = INDIC_INDIC_TXT_MAP[SelectedAWRPivotTableIndicator];
            FormatChartTitle(chart1, bt, st);


            // Relative chart
            ChartObject chartObj2 = sheet.ChartObjects().Add(50, chartTop + chartObj1.Height + 20, 600, 337);
            chartObj2.Name = AWR_RELATIVE_CHART_NAME;
            Excel.Chart chart2 = chartObj2.Chart;

            chart2.SetSourceData(AWRPivotTable.TableRange1, Type.Missing);
            chart2.ChartType = Excel.XlChartType.xlBarStacked100;
            chart2.ShowAllFieldButtons = false;

            // Format the chart
            chart2.HasTitle = true;

            bt = "Relative Use by Year";
            st = INDIC_INDIC_TXT_MAP[SelectedAWRPivotTableIndicator];
            FormatChartTitle(chart2, bt, st);

            // Set the colors of the AWR categories

            // Sort the seriesData list to enforce the order A → W → R → N**
            Dictionary<string, int> AwrOrders = new Dictionary<string, int>
            {
                { "A", 1 },
                { "W", 2 },
                { "R", 3 },
                { "N", 4 }
            };

            // Define colors
            Dictionary<string, int> AwrColors = new Dictionary<string, int>
            {
                { "A", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green) },      // Access: Green
                { "W", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange) },     // Watch: Light Orange
                { "R", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red) },        // Reserve: Red
                { "N", System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray) }        // Not classified / Not recommended: Grey
            };

            Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart1.SeriesCollection();
            foreach (Excel.Series series in seriesCollection)
            {
                var t = series.Name;
                // apply color based on category
                if (AwrColors.ContainsKey(series.Name))
                {
                    series.Format.Fill.ForeColor.RGB = AwrColors[series.Name];
                    //var order = AwrOrders[series.Name];
                    //var oldOrder = series.PlotOrder;
                    //series.PlotOrder = order; // get a crash when setting PlotOrder
                }
            }

            seriesCollection = (Excel.SeriesCollection)chart2.SeriesCollection();
            foreach (Excel.Series series in seriesCollection)
            {
                var t = series.Name;
                // apply color based on category
                if (AwrColors.ContainsKey(series.Name))
                {
                    series.Format.Fill.ForeColor.RGB = AwrColors[series.Name];
                    // series.PlotOrder = AwrOrders[series.Name];
                }
            }

            absoluteChart = chart1;
            relativeChart = chart2;
        }
    }
}
