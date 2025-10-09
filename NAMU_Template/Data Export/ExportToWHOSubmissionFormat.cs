// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using NAMU_Template.Helper;
using NAMU_Template.Models;
using static NAMU_Template.Helper.Constants;
using NAMU_Template.Data_Validation;
using NAMU_Template.Constants;
using AMU_Template.Validations;

namespace NAMU_Template.Data_Export
{
    public static class ExportToWHOSubmissionFormat
    {
        public static bool ValidateYearToExport(int yearToExport)
        {
            // Use SharedData.ProductConsummption directly
            var consProdData = SharedData.ProductConsummptionData ?? new List<ProductConsumption>();

            // Check if any record exists for the given year
            return consProdData.Any(p => p.Year == yearToExport);
        }
        //public static void ExportWHOGLASSAMCTemplatev1(int yearToExport, Dictionary<string, Product> prodData, Dictionary<int, object> consProdData)
        //{
        //    // Create a new workbook
        //    var excelApp = new Microsoft.Office.Interop.Excel.Application();
        //    var wb = excelApp.Workbooks.Add();
        //    wb.Title = "WHO GLASS AMC Submission Data";
        //    wb.Subject = "WHO GLASS AMC Submission Data Template";

        //    // Remove all sheets except the first
        //    excelApp.DisplayAlerts = false;
        //    int sheetCount = wb.Sheets.Count;
        //    for (int i = sheetCount - 1; i >= 1; i--)
        //    {
        //        var ws = (Worksheet)wb.Sheets[i];
        //        ws.Delete();
        //    }
        //    excelApp.DisplayAlerts = true;

        //    // Add "AMC Data" worksheet
        //    var wsAMCData = (Worksheet)wb.Worksheets.Add(Before: wb.Worksheets[wb.Worksheets.Count]);
        //    wsAMCData.Name = "AMC Data";

        //    // Create the product sheet header
        //    CreateProductSheetHeader(wsAMCData);

        //    // Export product data
        //    var prodLineDict = ExportProductData(prodData, yearToExport, wsAMCData);

        //    // Export consumption product data
        //    bool exported = ExportConsumptionProductData(consProdData, yearToExport, prodLineDict, wsAMCData);

        //    if (!exported)
        //    {
        //        System.Windows.Forms.MessageBox.Show("Error when exporting data to GLASS AMC format");
        //        return;
        //    }

        //    // Save as text/tab-delimited file
        //    var saveFileDialog = new SaveFileDialog
        //    {
        //        Filter = "GLASS Text Files (*.tsv)|*.tsv",
        //        Title = "Save GLASS AMC Submission Data"
        //    };

        //    if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //    {
        //        string tsvFilename = saveFileDialog.FileName;

        //        // Save the current system settings
        //        string oldDecSep = excelApp.DecimalSeparator;
        //        string old1000Sep = excelApp.ThousandsSeparator;
        //        bool oldSysSep = excelApp.UseSystemSeparators;

        //        // Set new separators for TSV format
        //        excelApp.DecimalSeparator = ".";
        //        excelApp.ThousandsSeparator = "";
        //        excelApp.UseSystemSeparators = false;

        //        // Save the workbook
        //        wb.SaveAs(tsvFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlTextWindows);

        //        // Restore the original system settings
        //        excelApp.DecimalSeparator = oldDecSep;
        //        excelApp.ThousandsSeparator = old1000Sep;
        //        excelApp.UseSystemSeparators = oldSysSep;
        //    }

        //    // Clean up
        //    wb.Close(false);
        //    excelApp.Quit();
        //}

        public static bool ExportNAMUDataToWIDP(int yearToExport, List<Product> prodData, List<ProductConsumption> productConsumptionData)
        {
            string widpFilePath;
            WIDPTemplateV1 widpTemplate = new WIDPTemplateV1();
            ErrorStatus errorStatus = new ErrorStatus();

            // Initialize the template
            widpTemplate.Initialize();

            // Get the WIDP template file path
            widpFilePath = GetWIDPNAMUTemplateFilePath();
            var (excelApp, widpWorkbook) = OpenWIDPNAMUTemplateFile(widpFilePath, widpTemplate, errorStatus);
            if (errorStatus.Status != EntityStatus.OK)
            {
                MessageBox.Show(errorStatus.ErrorsToString(), "Error opening the GLASS WIDP Template");
                return false;
            }
            // Make Excel visible for the user
            excelApp.Visible = true;

            ExportDataToWIDPWorkbook(widpWorkbook, yearToExport, prodData, productConsumptionData, widpTemplate);

            return true;

            // Clean up Excel application
            //widpWorkbook.Close(true); // Save changes
            //excelApp.Quit();
            //Marshal.ReleaseComObject(excelApp);
        }

        private static string GetWIDPNAMUTemplateFilePath()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Open GLASS NAMU Template File";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
                else
                {
                    return null;
                }
            }
        }
        public static (Microsoft.Office.Interop.Excel.Application, Workbook) OpenWIDPNAMUTemplateFile(string filePath, WIDPTemplateV1 widpTemplate, ErrorStatus errorStatus)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;

            try
            {
                //excelApp = Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                // Open the workbook
                workbook = excelApp.Workbooks.Open(filePath);

                bool wsRegFound = false;
                bool wsDataFound = false;

                foreach (Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == widpTemplate.GetRegisterWorksheetName())
                    {
                        wsRegFound = true;
                    }
                    else if (sheet.Name == widpTemplate.GetUseWorksheetName())
                    {
                        wsDataFound = true;
                    }

                    if (wsRegFound && wsDataFound)
                    {
                        break;
                    }
                }

                if (!wsRegFound || !wsDataFound)
                {
                    errorStatus.AddErrorMsgs("The file does not look like a GLASS-NAMC product level template file.");
                    workbook.Close(false);
                    return (excelApp, null);
                }

                return (excelApp, workbook);
            }
            catch (Exception ex)
            {
                errorStatus.AddErrorMsgs("Cannot open the file, it is not a valid Excel file. " + ex.Message);
                workbook?.Close(false);
                excelApp?.Quit();
                return (null, null);
            }
            //finally
            //{
            //    if (excelApp != null)
            //    {
            //        excelApp.Quit();
            //    }
            //}
        }

        public static void ExportDataToWIDPWorkbook(Workbook widpWorkbook, int yearToExport, List<Product> productData, List<ProductConsumption> consumptionProductData, WIDPTemplateV1 widpTemplate)
        {
            if (widpWorkbook == null || widpTemplate == null)
                throw new System.ArgumentNullException("Workbook or Template cannot be null.");

            // Fetch sheets collection
            Sheets sheets = null;
            Worksheet registerWorksheet = null;
            Worksheet useWorksheet = null;

            try
            {
                // Fetch sheets collection
                sheets = widpWorkbook.Sheets;

                // Fetch worksheet indices
                int registerIndex = widpTemplate.GetRegisterWorksheetIndex();
                int useIndex = widpTemplate.GetUseWorksheetIndex();
                string registerSheetName = widpTemplate.GetRegisterWorksheetName();
                string useSheetName = widpTemplate.GetUseWorksheetName();

                // Validate indices
                /*int sheetCount = sheets.Count; // Store count in a local variable
                if (sheetCount < registerIndex || sheetCount < useIndex)
                    throw new IndexOutOfRangeException("Worksheet index is out of range.");*/

                // Fetch worksheets
                registerWorksheet = sheets[registerSheetName] as Worksheet;
                useWorksheet = sheets[useSheetName] as Worksheet;

                if (registerWorksheet == null || useWorksheet == null)
                    throw new InvalidOperationException("One or more worksheets could not be found.");

                ExportRegisterDataToWorksheet(registerWorksheet, yearToExport, productData, widpTemplate);
                ExportUseDataToWorksheet(useWorksheet, yearToExport, productData, consumptionProductData, widpTemplate);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine($"COMException: {ex.Message}");
                throw;
            }
            //} finally
            //{
            //    // Release COM objects 
            //    if (registerWorksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(registerWorksheet);
            //    if (useWorksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(useWorksheet);
            //    if (sheets != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
            //}

        }

        private static void ExportRegisterDataToWorksheet(Worksheet ws, int yearToExport, List<Product> productData, WIDPTemplateV1 widpTemplate)
        {
            //Product pr;
            //var tmp = default(KeyValuePair<int, Product>);
            //int prK;
            int startLineNo;
            int lineNo;

            lineNo = widpTemplate.GetRegisterWorksheetStartRow();
            List<object[]> consolidatedData = new List<object[]>();

            foreach (var pr in productData)
            {

                if (!pr.IsProductStatusError())
                {
                    pr.Year = yearToExport;

                    /*object[] record = new object[26];
                    record[0] = pr.GetValueForVariable(Product.UID_FIELD);
                    record[1] = pr.GetValueForVariable(Product.COUNTRY_FIELD);
                    record[2] = pr.GetValueForVariable(Product.ENROLMENT_DATE_WIDP_FIELD);
                    record[3] = pr.GetValueForVariable(Product.ENROLMENT_DATE_WIDP_FIELD);
                    record[4] = pr.GetValueForVariable(Product.INCIDENT_DATE_WIDP_FIELD);
                    record[5] = pr.GetValueForVariable(Product.PRODUCT_ID_FIELD);
                    record[6] = pr.GetValueForVariable(Product.PRODUCT_NAME_FIELD);
                    record[7] = pr.GetValueForVariable(Product.LABEL_FIELD);
                    record[8] = FormatWIDP_Decimal((Decimal?)pr.GetValueForVariable(Product.PACKSIZE_FIELD), true);
                    record[9] = FormatWIDP_Decimal((Decimal?)pr.GetValueForVariable(Product.STRENGTH_FIELD), true);
                    record[10] = pr.GetValueForVariable(Product.STRENGTH_UNIT_FIELD);
                    record[11] = pr.GetValueForVariable(Product.STRENGTH_UNIT_FIELD);*/

                    consolidatedData.Add(new object[]
                    {
                        pr.GetValueForVariable(Product.UID_FIELD),
                        pr.GetValueForVariable(Product.COUNTRY_FIELD),
                        pr.GetValueForVariable(Product.ENROLMENT_DATE_WIDP_FIELD),
                        pr.GetValueForVariable(Product.ENROLMENT_DATE_WIDP_FIELD), // Repeat for formatting
                        pr.GetValueForVariable(Product.INCIDENT_DATE_WIDP_FIELD),
                        //pr.GetValueForVariable(VAR_PROD_INCID_DATE), // Repeat for formatting
                        pr.GetValueForVariable(Product.PRODUCT_ID_FIELD),
                        pr.GetValueForVariable(Product.PRODUCT_NAME_FIELD),
                        pr.GetValueForVariable(Product.LABEL_FIELD),
                        FormatWIDP_Decimal((Decimal?)pr.GetValueForVariable(Product.PACKSIZE_FIELD), true),
                        FormatWIDP_Decimal((Decimal?)pr.GetValueForVariable(Product.STRENGTH_FIELD), true),
                        pr.GetValueForVariable(Product.STRENGTH_UNIT_FIELD),
                        FormatWIDP_Decimal((Decimal?)pr.GetValueForVariable(Product.CONCENTRATION_VOLUME_FIELD), true),
                        FormatWIDP_Decimal((Decimal?)pr.GetValueForVariable(Product.VOLUME_FIELD), true),
                        pr.GetValueForVariable(Product.ATC5_FIELD),
                        pr.GetValueForVariable(Product.COMBINATION_FIELD),
                        pr.GetValueForVariable(Product.ROUTE_ADMIN_FIELD),
                        pr.GetValueForVariable(Product.SALT_FIELD),
                        pr.GetValueForVariable(Product.PAEDIATRIC_FIELD),
                        pr.GetValueForVariable(Product.FORM_FIELD),
                        pr.GetValueForVariable(Product.INGREDIENTS_FIELD),
                        pr.GetValueForVariable(Product.PRODUCT_ORIGIN_FIELD),
                        pr.GetValueForVariable(Product.MANUFACTURER_COUNTRY_FIELD),
                        pr.GetValueForVariable(Product.MARKET_AUTH_HOLDER_FIELD),
                        pr.GetValueForVariable(Product.GENERICS_FIELD),
                        FormatWIDP_Year((int?)pr.GetValueForVariable(Product.YEAR_AUTHORIZATION_FIELD), true),
                        FormatWIDP_Year((int?)pr.GetValueForVariable(Product.YEAR_WITHDRAWAL_FIELD), true),
                        VALUE_DATA_STATUS
                    });
                }
            }

            // Write data in batches
            int batchSize = 1000; // Adjust as needed
            int row = lineNo;
            for (int i = 0; i < consolidatedData.Count; i += batchSize)
            {
                var batch = consolidatedData.Skip(i).Take(batchSize).ToArray();
                var batch2D = ConvertTo2DArray(batch);

                var range = ws.Range[ws.Cells[row, 1], ws.Cells[row + batch2D.GetLength(0) - 1, batch2D.GetLength(1)]];
                range.Value = batch2D;

                row += batch2D.GetLength(0);
            }
        }
        private static object[,] ConvertTo2DArray(object[][] jaggedArray)
        {
            int rows = jaggedArray.Length;
            int cols = jaggedArray[0].Length;
            object[,] result = new object[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    result[i, j] = jaggedArray[i][j];
                }
            }

            return result;
        }

        private static void ExportUseDataToWorksheet(Worksheet ws, int yearToExport, List<Product> productData, List<ProductConsumption> consumptionProductData, WIDPTemplateV1 widpTemplate)
        {
            // Implementation for exporting use data goes here
            // You need to replace this with the actual logic from the corresponding VBA function
            int lineNo = widpTemplate.GetUseWorksheetStartRow();
            List<object[]> consolidatedData = new List<object[]>();
            int idx = 0;
            // Ensure consumptionProductData is not null
            consumptionProductData = SharedData.ProductConsummptionData ?? new List<ProductConsumption>();

            // Filter the list to only include records for the given year
            var filteredData = consumptionProductData.Where(p => p.Year == yearToExport).ToList();

            foreach (var consPack in filteredData)
            {
                /*var pr = productData.FirstOrDefault(p => p.ProductLineNo == consPack.LineNo);
                if (pr == null)
                {
                    continue; // Skip if product not found
                }*/

                string prUid = consPack.ProductUniqueId;

                if (consPack.AvailabilityTotal)
                {
                    consolidatedData.Add(new object[]
                    {
                        prUid,
                        null,
                        consPack.GetValueForVariable(ProductConsumption.PROD_CONS_EVENT_DATE_WIDP_FIELD),
                        HealthSectorLevelString.GetStringForHealthSector(consPack.Sector),
                        HealthSectorLevelString.GetStringForHealthLevel(HealthLevel.Total),
                        consPack.PKGConsumptionTotal,
                        VALUE_DATA_STATUS,
                    });
                }
                else
                {
                    if (consPack.AvailabilityCommunity)
                    {
                        consolidatedData.Add(new object[]
                        {
                            prUid,
                            null,
                            consPack.GetValueForVariable(ProductConsumption.PROD_CONS_EVENT_DATE_WIDP_FIELD),
                            HealthSectorLevelString.GetStringForHealthSector(consPack.Sector),
                            HealthSectorLevelString.GetStringForHealthLevel(HealthLevel.Community),
                            consPack.PKGConsumptionCommunity,
                            VALUE_DATA_STATUS,
                        });
                    }

                    if (consPack.AvailabilityHospital)
                    {
                        consolidatedData.Add(new object[]
                        {
                            prUid,
                            null,
                            consPack.GetValueForVariable(ProductConsumption.PROD_CONS_EVENT_DATE_WIDP_FIELD),
                            HealthSectorLevelString.GetStringForHealthSector(consPack.Sector),
                            HealthSectorLevelString.GetStringForHealthLevel(HealthLevel.Hospital),
                            consPack.PKGConsumptionHospital,
                            VALUE_DATA_STATUS,
                        });
                    }
                }
            }

            // Write data in batches
            int batchSize = 1000; // Adjust as needed
            int row = lineNo;
            for (int i = 0; i < consolidatedData.Count; i += batchSize)
            {
                var batch = consolidatedData.Skip(i).Take(batchSize).ToArray();
                var batch2D = ConvertTo2DArray(batch);

                var range = ws.Range[ws.Cells[row, widpTemplate.GetColumnIndexForUSEVariable(ProductConsumption.PROD_CONS_PROD_UID_FIELD)],
                                      ws.Cells[row + batch2D.GetLength(0) - 1, widpTemplate.GetColumnIndexForUSEVariable(ProductConsumption.PROD_CONS_STATUS_WIDP_FIELD)]];
                range.Value = batch2D;

                row += batch2D.GetLength(0);
            }

            //lineNo++;
            //if (idx % 100 == 0)
            //{
            //    // Simulates DoEvents in VBA
            //    System.Windows.Forms.Application.DoEvents();
            //}
            //idx++;
        }

        private static Decimal? FormatWIDP_Decimal(Decimal? val, bool excludeZero=true)
        {
            if (val == null) { return null; }
            if (excludeZero && val == Decimal.Zero) { return null; }
            return val;
        }

        private static int? FormatWIDP_Year(int? val, bool excludeZero = true)
        {
            if (val == null) { return null; }
            if (excludeZero && val == 0) { return null; }
            return val;
        }

        private static string FormatWIDP_HLevel(HealthLevel inLevel)
        {
            string outLevel = string.Empty;
            switch (inLevel)
            {
                case HealthLevel.Total:
                    outLevel = VALUE_H_LEVEL_TOTAL;
                    break;
                case HealthLevel.Community:
                    outLevel = VALUE_H_LEVEL_COMMUNITY;
                    break;
                case HealthLevel.Hospital:
                    outLevel = VALUE_H_LEVEL_HOSPITAL;
                    break;
            }
            return outLevel;
        }
    }
}
