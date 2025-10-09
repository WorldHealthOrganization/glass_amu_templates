using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using NAMU_Template.Constants;
using NAMU_Template.Data_Validation;
using NAMU_Template.Models;
using NAMU_Template.Data_Export;
using NAMU_Template.Data_Parsing;
using NAMU_Template.Data_Processing;
using NAMU_Template.Helper;
using Excel = Microsoft.Office.Interop.Excel;
using AMU_Template.Constants;
using AMU_Template.Validations;

namespace NAMU_Template
{

    public static class CustomLoadingBox
    {
        private static Form _loadingForm;

        public static void Show(string message, string title = "Please wait")
        {
            if (_loadingForm != null)
            {
                Close(); // Ensure no duplicate forms
            }

            _loadingForm = new Form
            {
                Text = title,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false,
                ClientSize = new System.Drawing.Size(300, 100),
                ControlBox = false // Disable close button
            };

            Label label = new Label
            {
                Text = message,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };

            _loadingForm.Controls.Add(label);

            // Set Excel window as the owner of this form
            var excelHandle = (IntPtr)Globals.ThisWorkbook.Application.Hwnd;
            SetParent(_loadingForm.Handle, excelHandle);

            _loadingForm.Show();

            // Force the UI to refresh
            System.Windows.Forms.Application.DoEvents();
        }

        public static void Close()
        {
            if (_loadingForm != null)
            {
                _loadingForm.Close();
                _loadingForm.Dispose();
                _loadingForm = null;
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
    }


    public static class CustomInputBox
    {
        public static string Show(string prompt, string title, string defaultValue = "")
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = prompt;
            textBox.Text = defaultValue;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new System.Drawing.Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            return dialogResult == DialogResult.OK ? textBox.Text : null;
        }
    }

    public partial class Feuil1
    {

        Excel.Application excelApp = new Excel.Application();

        //Booolean check values for parsing the data..!
        bool parseCheckAvailabilityData;

        //Boolean check values for validation of the data..!
        bool checkValidation = false;
        bool isProductValidating;
        public Excel.Range StatusCell => this.Cells[4, 4];
        public static int status_value = 0;
        private const int VSTATUS_PARSED = 1;
        private const int VSTATUS_CALC = 2;
        private const int VSTATUS_EXPORT = 4;
        private const string DATA_SHEET = "DataSheet";

        // public readonly VStatus VSTATUS;

        List<DataAvailability> availData;
        List<Population> popYears;

        private void Feuil1_Startup(object sender, System.EventArgs e)
        {
            Excel.Range cellC2 = this.Range["C2"];
            cellC2.Value = "ATC/DDD index";

            Excel.Range cellE2 = this.Range["E2"];
            cellE2.Value = "2025v1";

            Excel.Range cellC4 = this.Range["C4"];
            cellC4.Value = "Status:";

            Excel.Range cellD4 = this.Range["D4"];
            cellD4.Value = "NA"; // Dynamic value, will be updated based on button clicks

            foreach (Excel.Shape shape in Globals.Feuil1.Shapes)
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoFormControl)
                {
                    System.Diagnostics.Debug.WriteLine($"Deleting Shape: {shape.Name}");
                    shape.Delete();
                }
            }

            //Control Buttons..!
            var validateButton = this.Controls.AddButton(Range["C9", "E11"], "validate");
            validateButton.Text = "Validate Products";
            validateButton.Border.Color = Color.Black;
            validateButton.Click += ValidateProducts_Click;

            var calculateUseButton = this.Controls.AddButton(Range["G9", "I11"], "calculateUse");
            calculateUseButton.Text = "Calculate Use";
            calculateUseButton.Border.Color = Color.Black;
            calculateUseButton.Click += CalculateUse_Click;

            var exportCalculatedUseDataButton = this.Controls.AddButton(Range["C15", "E17"], "exportCalculatedUseData");
            exportCalculatedUseDataButton.Text = "Export Calculated Use Data";
            exportCalculatedUseDataButton.Border.Color = Color.Black;
            exportCalculatedUseDataButton.Click += ExportCalculatedUseData_Click;

            var exportForWHOGlassButton = this.Controls.AddButton(Range["G15", "I17"], "exportForWHOGlass");
            exportForWHOGlassButton.Text = "Export for WHO Glass AMU Submission";
            exportForWHOGlassButton.Border.Color = Color.Black;
            exportForWHOGlassButton.Click += ExportForWHOGlass_Click;
        }

        private void ResetExcelSheet(Excel.Worksheet sheet)
        {

            // Check if the sheet is not null..!
            if (sheet != null)
            {
                // Resetting borders for the entire sheet..!
                Excel.Range usedRange = sheet.UsedRange;

                //any previous data..!
                Excel.Range dataRange = usedRange[3, usedRange.Rows.Count];

                if (usedRange != null)
                {
                    dataRange.ClearContents();
                    sheet.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
            }
        }

        private void CalculateUse_Click(object sender, EventArgs e)
        {
            try
            {
                //MessageBox.Show("Calculate logic triggered!");
                if (Validator.GetStatus() < VStatus.PARSED)
                {
                    MessageBox.Show("Please validate products first", "Validation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                // Show the loading dialog
                CustomLoadingBox.Show("Calculation in progress, please wait...");

                bool success = PerformCalculation();
                if (!success)
                {
                    return;
                }
                Validator.SetStatus(VStatus.CALCULATED);
            }
            catch (Exception ex)
            {
                // Handle exceptions
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Close the loading dialog
                CustomLoadingBox.Close();
            }

        }
        private bool PerformCalculation()
        {
            try
            {
                // Use cached data if available, otherwise parse
                if (SharedData.AvailData == null || !SharedData.AvailData.Any())
                {
                    Excel.Worksheet availabilityDataSheet = LoadWorkSheet("Availability Data");
                    if (!TryParseAvailabilityData(availabilityDataSheet, out List<DataAvailability> availData))
                    {
                        MessageBox.Show("Availability data parsing failed or returned no data.");
                        return false;
                    }
                    SharedData.AvailData = availData;
                }

                if (SharedData.PopYears == null || !SharedData.PopYears.Any())
                {
                    Excel.Worksheet populationDataSheet = LoadWorkSheet("Population Data");
                    if (!TryParsePopulationData(populationDataSheet, SharedData.AvailData, out List<Population> popYears))
                    {
                        MessageBox.Show("No population data have been provided.");
                        return false;
                    }
                    SharedData.PopYears = popYears;
                }

                bool validAP = ValidatePopulationAvailability(SharedData.AvailData, SharedData.PopYears);
                if (!validAP)
                {
                    return false;
                }

                Excel.Worksheet dataSheet = (Excel.Worksheet)this.Application.Worksheets[TemplateFormat.DATA_SHEETNAME];
                int[] years = ProductDataParser.ParsePackYears(dataSheet);

                if (years.Length < 0)
                {
                    MessageBox.Show(
                        "The number of years of data is not valid. Each year should be repeated three times: Total sector, Community sector, and Hospital sector even if you are not providing data for all three sectors.",
                        "Invalid Data",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return false;
                }

                var productData = SharedData.Products;
                var es = new ErrorStatus(); // Create ErrorStatus here
                var productConsumptionData = ProductDataParser.ParsePackages(years, dataSheet, productData, SharedData.PopYears, SharedData.AvailData, es);
                ConsumptionProcessing.CalculateDDDConsumption(productConsumptionData, years);
                SharedData.ProductConsummptionData = productConsumptionData;
                // Show errors if any exist
                if (es.Errors.Any())
                {
                    MessageBox.Show(string.Join("\n", es.Errors.Distinct()));
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
                return false;
            }
        }

        private void ExportCalculatedUseData_Click(object sender, EventArgs e)
        {
            try
            {
                if (Validator.GetStatus() != VStatus.CALCULATED)
                {
                    MessageBox.Show("Please validate and calculate products first", "Validation, Calculate Use Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                // Show the loading dialog
                CustomLoadingBox.Show("Export in progress, please wait...");

                // Execute your logic asynchronously
                if (!PerformCalculation()) // Call the extracted calculation method
                {
                    return;
                }
                ExportConsumption.ExportCalculateUseConsumption(availData);
                Validator.SetStatus(VStatus.EXPORTED);
            }
            catch (Exception ex)
            {
                // Handle exceptions
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Close the loading dialog
                CustomLoadingBox.Close();
            }
        }

        private void ExportForWHOGlass_Click(object sender, EventArgs e)
        {
            // Check if productConsumptionData is available
            if (SharedData.ProductConsummptionData == null) // might be null
            {
                MessageBox.Show("Please validate and calculate before exporting the data.");
                return;
            }
            //string input = String.Empty;
            // Prompt the user to select the year to export  this need to take a look
            //string input = Microsoft.VisualBasic.InputBox(
            //    "Select the year of data to be exported...",
            //    "Year selection",
            //    "0");
            // Prompt the user to select the year to export
            string input = CustomInputBox.Show(
                            "Select the year of data to be exported...",
                            "Year selection",
                            "0");

            if (!int.TryParse(input, out int yearToExport) || yearToExport == 0)
            {
                return; // Exit if input is invalid or 0
            }

            // Validate the year
            if (!ExportToWHOSubmissionFormat.ValidateYearToExport(yearToExport))
            {
                System.Windows.Forms.MessageBox.Show($"The year {yearToExport} is not valid.");
                return;
            }

            // Export data to WIDP format
            bool ret = ExportToWHOSubmissionFormat.ExportNAMUDataToWIDP(yearToExport, SharedData.Products, SharedData.ProductConsummptionData);
            if (ret)
            {
                MessageBox.Show("You can save the GLASS WIDP file with a different name, close it. You can submit this new file on the GLASS platform.", "Export Successful...");
            }
        }

        private bool TryParsePopulationData(Excel.Worksheet populationDataSheet, List<DataAvailability> availData, out List<Population> popYears)
        {
            popYears = new List<Population>();
            var usedRange = populationDataSheet.UsedRange;
            int noOfPops = NAMU_Template.Helper.Utils.GetRowsCountPopulationData(usedRange);
            int firstRow = usedRange.Row + 1;
            int maxRowsForPopulation = firstRow + noOfPops - 1;
            if (maxRowsForPopulation < 2)
            {
                popYears = null;
                return false;
            }
            bool isParsed = AvailabilityPopulationDataParser.ParsePopulation(maxRowsForPopulation, firstRow, usedRange, out popYears, availData);

            if (isParsed && popYears != null && popYears.Any())
            {
                SharedData.PopYears = popYears; // Store parsed data
                return true;
            }

            return false;
        }

        private bool TryParseAvailabilityData(Excel.Worksheet availabilityDataSheet, out List<DataAvailability> availData)
        {
            availData = new List<DataAvailability>();
            var usedRange = availabilityDataSheet.UsedRange;
            //int rowWithData = Helper.Utils.GetRowsCountAvailabilityData(usedRange);
            int[] years = AvailabilityPopulationDataParser.ParseAvailabilityYears(usedRange.Rows[5]); // Assuming the years are in row 5
            bool isParsed = AvailabilityPopulationDataParser.ParseAvailability(availabilityDataSheet, years, out availData);

            if (isParsed && availData != null && availData.Any())
            {
                SharedData.AvailData = availData; // Store parsed data
                return true;
            }
            return false;
        }


        private void ValidateProducts_Click(object sender, EventArgs e)
        {
            try
            {
                // Show the loading dialog
                CustomLoadingBox.Show("Validation in progress, please wait...");
                Excel.Worksheet productDataSheet = LoadWorkSheet(TemplateFormat.PRODUCT_SHEETNAME);

                // Use cached data if available, otherwise parse
                if (SharedData.AvailData == null || !SharedData.AvailData.Any())
                {
                    Excel.Worksheet availabilityDataSheet = LoadWorkSheet(TemplateFormat.AVAILABILITY_SHEETNAME);
                    if (!TryParseAvailabilityData(availabilityDataSheet, out availData))
                    {
                        MessageBox.Show("Availability data parsing failed or returned no data. Please check the input.");
                        return;
                    }
                }

                if (SharedData.PopYears == null || !SharedData.PopYears.Any())
                {
                    Excel.Worksheet populationDataSheet = LoadWorkSheet(TemplateFormat.POPULATION_SHEETNAME);
                    if (!TryParsePopulationData(populationDataSheet, SharedData.AvailData, out popYears))
                    {
                        MessageBox.Show("No population data have been provided.");
                        return;
                    }
                }

                List<Product> listProducts = ProductDataParser.ParseProducts();
                // Populate products
                SharedData.Products = listProducts;

                // ProductValidator productValidator = new ProductValidator();


                foreach (var pr in listProducts)
                {
                    //Call the validateProduct method
                    // pr.ValidateProduct(pr, true); OLD CODE
                    pr.ValidateProduct(false);

                    if (pr.IsProductValid())
                    {
                        if (pr.ATC5.Code != AMUConstants.ATC_Z99_CODE)
                        {
                            if (!ProductDataParser.IsATCClassInAvailability(pr.ATCClass, availData))
                            {
                                pr.AddWarningMsg("The ATC code is not among the antimicrobial classes specified in the availability data. It will be excluded from the calculation/export.");
                                pr.SetValidate(true, Product.ATCCLASS_VALIDATION);
                            }
                        }
                    }
                    else
                    {
                        if (pr.GetValidate(Product.ATC5_VALIDATION) && pr.ATC5.Code != AMUConstants.ATC_Z99_CODE)
                        {
                            if (!ProductDataParser.IsATCClassInAvailability(pr.ATCClass, availData))
                            {
                                pr.AddWarningMsg("The ATC code is not among the antimicrobial classes specified in the availability data. It will be excluded from the calculation/export.");
                                pr.SetValidate(true, Product.ATCCLASS_VALIDATION);
                            }
                        }
                    }
                }

                // Call the displayProduct method
                DisplayProductsAndStatuses(listProducts);
                ////Set isProductValidating to false
                isProductValidating = false;
                Validator.SetStatus(VStatus.PARSED);

            }
            catch (Exception ex)
            {
                // Handle exceptions
                MessageBox.Show($"An unexpected error occurred: {ex.Message}");
            }
            finally
            {
                // Close the loading dialog
                CustomLoadingBox.Close();
            }

        }

        public bool ValidatePopulationAvailability(List<DataAvailability> dataAvailability, List<Population> populationData)
        {
            // Create availability keys (year, AMClass, sector) using LINQ
            var availabilityKeys = dataAvailability.Select(data => data.Key).ToHashSet();

            var populationKeys = populationData
                .Select(data => data.Key)
                .ToHashSet();

            // Check for missing keys in population data
            foreach (var key in availabilityKeys)
            {
                if (!availabilityKeys.Contains(key))
                {
                    // Show an error message and return false
                    MessageBox.Show(
                       $"Population data are missing for year {key.Year}, AM class {key.AMClass}, and sector {key.Sector}. Please provide population data accordingly to availability data or correct availability.",
                       "Missing Population Data",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Warning
                    );

                    return false;
                }
            }

            foreach (var key in populationKeys)
            {
                if (!populationKeys.Contains(key))
                {
                    // Show an error message and return false
                    MessageBox.Show(
                       $"Population data are missing for year {key.Year}, AM class {key.AMClass}, and sector {key.Sector}. Please provide population data accordingly to availability data or correct availability.",
                       "Missing Population Data",
                       MessageBoxButtons.OK,
                       MessageBoxIcon.Warning
                    );

                    return false;
                }
            }

            // All keys match, return true
            return true;
        }

        public Dictionary<string, List<string>> AvailabilitiesYearATClassSector(Dictionary<int, Dictionary<string, Dictionary<string, DataAvailability>>> availData)
        {
            var yass = new Dictionary<string, List<string>>();

            foreach (var yearEntry in availData)
            {
                int year = yearEntry.Key;
                var yAvail = yearEntry.Value;

                if (yAvail.Count > 0)
                {
                    foreach (var amClassEntry in yAvail)
                    {
                        string amClass = amClassEntry.Key;
                        var amAvail = amClassEntry.Value;

                        foreach (var sectorEntry in amAvail)
                        {
                            var avail = sectorEntry.Value;

                            var yas = new List<string>
                        {
                            year.ToString(),
                            amClass,
                            HealthSectorLevelString.GetStringForHealthSector(avail.Sector),
                            avail.AvailabilityTotal ? "T" : "F",
                            avail.AvailabilityCommunity ? "T" : "F",
                            avail.AvailabilityHospital ? "T" : "F"
                        };

                            string key2 = $"{yas[0]}{yas[1]}{yas[2]}|{yas[3]}{yas[4]}{yas[5]}";
                            string availLvs = $"{yas[3]}{yas[4]}{yas[5]}";

                            if (availLvs != "FFF")
                            {
                                yass[key2] = yas;
                            }
                        }
                    }
                }
            }

            return yass;
        }
        public Dictionary<string, List<string>> PopulationsYearAMClassSector(Dictionary<int, Dictionary<string, Dictionary<string, Population>>> popData)
        {
            // Predefined ATC codes
            string[] atcs = { "A07AA", "D01BA", "J01", "J02", "J04", "J05", "P01AB", "P01B" };

            // Dictionary to hold the result
            var yass = new Dictionary<string, List<string>>();

            // Iterate over years
            foreach (var yearEntry in popData)
            {
                int year = yearEntry.Key;
                string yearStr = year.ToString();
                var yPop = yearEntry.Value;

                // Check if the year has data
                if (yPop.Count > 0)
                {
                    // Iterate over ATC classes
                    foreach (var amClassEntry in yPop)
                    {
                        string amClass = amClassEntry.Key;
                        var amPop = amClassEntry.Value;

                        // If the class is "ALL", process for all predefined ATC codes
                        if (amClass == "ALL")
                        {
                            foreach (var atc in atcs)
                            {
                                GetPopAvailAmClass(yearStr, atc, amPop, yass);
                            }
                        }
                        else
                        {
                            GetPopAvailAmClass(yearStr, amClass, amPop, yass);
                        }
                    }
                }
            }

            return yass;
        }
        private void GetPopAvailAmClass(string year, string amClass, Dictionary<string, Population> amPop, Dictionary<string, List<string>> yass)
        {
            // Check if there are keys in the amPop dictionary
            if (amPop.Count > 0)
            {
                // Iterate through each sector in the amPop dictionary
                foreach (var sector in amPop.Keys)
                {
                    // Get the population object for the current sector
                    var pop = amPop[sector];

                    // Check the sector and add availability levels accordingly
                    if (sector == "GLO")
                    {
                        GetPopAvailLevelAmClass(year, amClass, "GLO", pop, yass);
                        GetPopAvailLevelAmClass(year, amClass, "PUB", pop, yass);
                        GetPopAvailLevelAmClass(year, amClass, "PRI", pop, yass);
                    }
                    else
                    {
                        GetPopAvailLevelAmClass(year, amClass, sector, pop, yass);
                    }
                }
            }
        }
        private void GetPopAvailLevelAmClass(string year, string amClass, string sector, Population pop, Dictionary<string, List<string>> yass)
        {
            // Create an array to store the availability details
            var yas = new List<string>(6)
            {
                year,                                       // Year
                amClass,                                    // ATC class
                sector,                                     // Sector
                pop.TotalPopulation > 0 ? "T" : "F",        // Total Population
                pop.CommunityPopulation > 0 ? "T" : "F",    // Community Population
                pop.HospitalPopulation > 0 ? "T" : "F"      // Hospital Population 
            };

            // Construct the key
            string key2 = $"{yas[0]}{yas[1]}{yas[2]}|{yas[3]}{yas[4]}{yas[5]}";

            // Add the key-value pair to the dictionary
            yass[key2] = yas;
        }

        private Excel.Worksheet LoadWorkSheet(string sheetName)
        {
            return this.Application.ActiveWorkbook.Worksheets[sheetName];
        }

        private void DisplayProductsAndStatuses(List<Product> listProducts) // TODO: To be moved out of the Product class
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.ActiveWorkbook;
            Excel.Worksheet productWorksheet = wb.Worksheets[TemplateFormat.DATA_SHEETNAME];
            Color InfoColor = Color.FromArgb(255, 220, 238, 157);

            //Logic:
            //Instead of interacting with the excel object for every product, create a data object and populate that, after populating the obj, assign it to the Excel Range..!
            int totalRows = listProducts.Count;
            int totalColumns = 9; // Assuming there are 7 columns to display
            int dataStartColumn = TemplateFormat.AUTO_CALC_START_COL_IDX;
            int dataEndColumn = dataStartColumn + TemplateFormat.AUTO_CALC_END_COL_IDX;

            if (totalRows == 0)
            {
                MessageBox.Show($"No product to validate.\r\n");
                return;
            }

            // Create a 2D array for data..!
            //For Rows and Columns..!
            object[,] data = new object[totalRows, totalColumns];
            // Create a 2D array for status data..!
            object[,] statusData = new object[totalRows, 2];
            bool hasErrors = false;

            // Clear the existing formatting and values for the status columns
            Excel.Range statusRange = productWorksheet.Range[productWorksheet.Cells[listProducts[0].ProductLineNo, 1],
                                                             productWorksheet.Cells[listProducts[totalRows - 1].ProductLineNo, 2]];
            statusRange.Interior.Color = Color.White; // Reset cell background color to white
            statusRange.Value2 = ""; // Clear existing messages

            for (int i = 0; i < totalRows; i++)
            {
                Product pr = listProducts[i];
                int lineNo = pr.ProductLineNo;

                // Populate data array
                if (pr.GetValidate(Product.CONVERSION_VALIDATION) && pr.ConversionFactor != 1)
                {
                    data[i, TemplateFormat.AUTO_CALC_CONVERSION_FACTOR_COL_IDX] = pr.ConversionFactor;
                }
                if (pr.GetValidate(Product.CONTENT_VALIDATION))
                {
                    data[i, TemplateFormat.AUTO_CALC_CONTENT_COL_IDX] = pr.Content.Value;
                    data[i, TemplateFormat.AUTO_CALC_CONTENT_UNIT_COL_IDX] = pr.Content.Unit.Code;
                }
                if(pr.GetValidate(Product.ARS_VALIDATION))
                {
                    data[i, TemplateFormat.AUTO_CALC_ARS_COL_IDX] = pr.ARS;
                }
                if (pr.GetValidate(Product.DDD_VALIDATION))
                {
                    data[i, TemplateFormat.AUTO_CALC_DDD_VALUE_COL_IDX] = pr.DDD.Value;
                    data[i, TemplateFormat.AUTO_CALC_DDD_UNIT_COL_IDX] = pr.DDD.Unit.Code;
                    
                }
                if(pr.GetValidate(Product.CONVERSION_VALIDATION) && pr.GetValidate(Product.CONTENT_VALIDATION) && pr.GetValidate(Product.DDD_VALIDATION))
                {
                    data[i, TemplateFormat.AUTO_CALC_DPP_COL_IDX] = pr.NbDDD;
                }
                    
                // Safe null checks for pr.AWaRe and pr.MEML
                data[i, TemplateFormat.AUTO_CALC_AWR_COL_IDX] = pr.AWaRe ?? "N/A"; // Default to "N/A" if pr.AWaRe is null
                data[i, TemplateFormat.AUTO_CALC_MEML_COL_IDX] = YesNoNAString.GetStringFromYesNoNA(pr.MEML);

                // Populate statusData array..!
                switch (pr.Status)
                {
                    case EntityStatus.OK:
                        statusData[i, TemplateFormat.DATA_SHEET_STATUS_COL_IDX-1] = "OK";
                        statusData[i, TemplateFormat.DATA_SHEET_STATUS_MSG_COL_IDX - 1] = "";
                        productWorksheet.Range[productWorksheet.Cells[lineNo, 1], productWorksheet.Cells[lineNo, 2]].Interior.Color = Color.Green;
                        productWorksheet.Range[productWorksheet.Cells[lineNo, 1], productWorksheet.Cells[lineNo, 2]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        break;
                    case EntityStatus.INFO:
                        statusData[i, 0] = "INFO";
                        statusData[i, 1] = pr.GetStatusMessages();
                        productWorksheet.Range[productWorksheet.Cells[lineNo, 1], productWorksheet.Cells[lineNo, 2]].Interior.Color = InfoColor;// Color.YellowGreen;
                        productWorksheet.Range[productWorksheet.Cells[lineNo, 1], productWorksheet.Cells[lineNo, 2]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        break;
                    case EntityStatus.WARNING:
                        statusData[i, 0] = "WARNING";
                        statusData[i, 1] = pr.GetStatusMessages();
                        productWorksheet.Range[productWorksheet.Cells[lineNo, 1], productWorksheet.Cells[lineNo, 2]].Interior.Color = Color.Orange;
                        productWorksheet.Range[productWorksheet.Cells[lineNo, 1], productWorksheet.Cells[lineNo, 2]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        break;
                    case EntityStatus.ERROR:
                        string statusText = "ERROR";
                        string msgTxt = pr.GetStatusMessages();
                        statusData[i, 0] = statusText;
                        statusData[i, 1] = msgTxt;
                        productWorksheet.Range[productWorksheet.Cells[lineNo, 1], productWorksheet.Cells[lineNo, 2]].Interior.Color = Color.Red;
                        productWorksheet.Range[productWorksheet.Cells[lineNo, 1], productWorksheet.Cells[lineNo, 2]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        foreach (var validation in pr.ValidationMessages)
                        {
                            hasErrors = true;
                            dynamic dataField = validation.ErrorField;
                            if (dataField != null)
                            {
                                int columnIndex = dataField.FieldColumn;
                                Excel.Range cellWithIssue = productWorksheet.Cells[lineNo, columnIndex];
                                cellWithIssue.Borders.Color = Color.Red;
                                cellWithIssue.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            }
                        }

                        // Set the concatenated validation messages..!
                        productWorksheet.Cells[lineNo, 2].Value = msgTxt;
                        break;
                }
            }


            // Set the entire range at once for both data and status
            productWorksheet.Range[productWorksheet.Cells[listProducts[0].ProductLineNo, dataStartColumn], productWorksheet.Cells[listProducts[totalRows - 1].ProductLineNo, dataEndColumn]].Value2 = data;
            productWorksheet.Range[productWorksheet.Cells[listProducts[0].ProductLineNo, 1], productWorksheet.Cells[listProducts[totalRows - 1].ProductLineNo, 2]].Value2 = statusData;

            if (hasErrors)
            {
                // Show products with errors in the MessageBox..!
                MessageBox.Show($"Validation of the products has been completed.\n\nThere are some errors in the Product data sheet.\n\nPlease review the data and validate it again.");
            }
            else
            {
                MessageBox.Show($"Validation of the products has been successfully completed.\n\nPlease proceed by pressing the Calculate Use button to calculate the use in DDD.\r\n");
            }
            // Autofit columns after the loop
            productWorksheet.Columns.AutoFit();
            productWorksheet.Application.ActiveWorkbook.Save();
        }

        private void Feuil1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Feuil1_Startup);
            this.Shutdown += new System.EventHandler(Feuil1_Shutdown);
        }

        #endregion

    }
}
