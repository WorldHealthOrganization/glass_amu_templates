using AMU_Template.Constants;
using AMU_Template.Validations;
using HAMU_Template.Constants;
using HAMU_Template.Data_Export;
using HAMU_Template.Data_Parsing;
using HAMU_Template.Data_Processing;
using HAMU_Template.Data_Validation;
using HAMU_Template.Helper;
using HAMU_Template.Models;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace HAMU_Template
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

    public partial class Sheet3
    {

        //Booolean check values for parsing the data..!
        bool parseCheckAvailabilityData;
        bool parseCheckHospitalStructure;
        bool parseCheckActivityData;

        //Boolean check values for validation of the data..!
        bool checkValidation = false;

        // Get Open Excel
        Excel.Application excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

        public Excel.Range StatusCell;
        public static int optionValue = 0;
        public static int OPTION_PRODUCTS = 1;
        public static int OPTION_SUBSTANCES = 2;
     
        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
            //Excel.Range cellC2 = this.Range["C2"];
            //cellC2.Value = "ATC/DDD index";

            //Excel.Range cellC4 = this.Range["C4"];
            //cellC4.Value = "Status: ";

            //Excel.Range cellC7 = this.Range["C7"];
            //cellC7.Value = "Type of consumption data";
            //cellC7.Font.Bold = true;
            //cellC7.Font.Size = 14;

            //Excel.Range cellF2 = this.Range["F2"];
            //cellF2.Value = "2023";

            //Excel.Range cellA1 = this.Range["A1"];

            // Create the radio buttons..!
            RadioButton btnTypeProducts = new RadioButton
            {
                Name = "Option1",
                Text = "Products"
            };

            RadioButton btnTypeSubstances = new RadioButton
            {
                Name = "Option2",
                Text = "Substances"
            };

            //Add them to a GroupBox..!
            var dataTypeGrp = new System.Windows.Forms.GroupBox();

            //Set thier visibility location in the group box..!
            btnTypeProducts.Location = new System.Drawing.Point(0, 0);
            btnTypeSubstances.Location = new System.Drawing.Point(130, 0);
            dataTypeGrp.Name = "Options";
            dataTypeGrp.Text = "Type of data";

            dataTypeGrp.Controls.Add(btnTypeProducts);
            dataTypeGrp.Controls.Add(btnTypeSubstances);

            this.Controls.AddControl(dataTypeGrp, Range["E7", "I9"], "GroupOptions");

            btnTypeProducts.Click += DatatType_Click;
            btnTypeSubstances.Click += DatatType_Click;

            //Control Buttons..!
            var validateButton = this.Controls.AddButton(Range["C11", "E13"], "validate");
            validateButton.Text = "Validate Consumption";
            validateButton.Border.Color = Color.Black;
            validateButton.Click += Validate_Click;

            var calculateButton = this.Controls.AddButton(Range["G11", "I13"], "calculate");
            calculateButton.Text = "Calculate Consumption";
            calculateButton.Border.Color = Color.Black;
            calculateButton.Click += Calculate_Click;

            Excel.Worksheet substanceDataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.SUBSTANCE_SHEETNAME];
            ResetExcelSheet(substanceDataSheet);
            Excel.Worksheet productDataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.PRODUCT_SHEETNAME];
            ResetExcelSheet(productDataSheet);

            // cellA1.Select();
        }

        private Excel.Worksheet LoadWorkSheet(string sheetName)
        {
            return this.Application.ActiveWorkbook.Worksheets[sheetName];
        }

        private void DatatType_Click(object sender, EventArgs e)
        {
            var selectedRadio = sender as RadioButton;
            StatusCell = Cells.Range["D4"];
            int newValue;
            if (selectedRadio.Text == "Products")
            {
                newValue = OPTION_PRODUCTS;
            }
            else
            {
                newValue = OPTION_SUBSTANCES;
            }
            if (newValue==optionValue)
            {
                return;
            }
            optionValue = newValue;
            checkValidation = false;
        }

        private void Calculate_Click(object sender, EventArgs e)
        {
            if (optionValue == 0)
            {
                MessageBox.Show("Please select an appropriate option.");
                return;
            }

            if (!checkValidation)
            {
                MessageBox.Show("Please validate the data first");
                return;
            }

            if (optionValue == 1)
            {
                CalculateAndExportProductConsumption();
            }
            else
            {
                CalculateAndExportSubstanceConsumption();
            }

        }

        private bool CalculateAndExportProductConsumption()
        {
            List<ProductConsumption> consumptionData = new List<ProductConsumption>();

            

            Excel.Worksheet dataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.PRODUCT_SHEETNAME];
            int[] years = ProductDataParser.ParseConsYears(dataSheet);
            if (years.Count()==0)
            {

                MessageBox.Show(
                        "The number of years of data is not valid in the Product data worksheet.",
                        "Invalid Data",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

              return false;
            }


            ErrorStatus es = new ErrorStatus();
            consumptionData = ProductDataParser.ParseCons(years, dataSheet, ThisWorkbook.ProductData, ThisWorkbook.AvailabilityData, ThisWorkbook.HospitalActivityData, ThisWorkbook.HospitalStructureData, es);

            if (es.Errors.Any())
            {
                MessageBox.Show(string.Join("\n", es.Errors.Distinct()));
                return false;
            }

            if (consumptionData != null)
            {
                List<AtcConsumption> calculatedDDDConsumptions = ConsumptionProcessing.CalculateAtcDDDConsumption(consumptionData.Cast<MedicineConsumption>().ToList());
                Cursor.Current = Cursors.WaitCursor;
                ThisWorkbook.VSTATUS = VStatus.CALCULATED;
                var consExporter = new ExporterConsumption();
                consExporter.ExportAtcConsumption(calculatedDDDConsumptions);
                Cursor.Current = Cursors.Default;
                return true;
            }
            else
            {
                MessageBox.Show("There are some issues in the consumption data. Cannot proceed further.\nPlease enter correct data and Validate again.", 
                    "No data calculated", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning);
                ThisWorkbook.VSTATUS = VStatus.DIRTY;
                return false;
            }

            
        }

        private bool CalculateAndExportSubstanceConsumption()
        {
            List<SubstanceConsumption> consumptionData = new List<SubstanceConsumption>();



            Excel.Worksheet dataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.SUBSTANCE_SHEETNAME];
            int[] years = SubstanceDataParser.ParseConsYears(dataSheet);
            if (years.Count() == 0)
            {

                MessageBox.Show(
                        "The number of years of data is not valid in the Substance data worksheet.",
                        "Invalid Data",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                return false;
            }


            ErrorStatus es = new ErrorStatus();
            consumptionData = SubstanceDataParser.ParseCons(years, dataSheet, ThisWorkbook.SubstanceData, ThisWorkbook.AvailabilityData, ThisWorkbook.HospitalActivityData, ThisWorkbook.HospitalStructureData, es);

            if (es.Errors.Any())
            {
                MessageBox.Show(string.Join("\n", es.Errors.Distinct()));
                return false;
            }

            if (consumptionData != null)
            {
                List<AtcConsumption> calculatedDDDConsumptions = ConsumptionProcessing.CalculateAtcDDDConsumption(consumptionData.Cast<MedicineConsumption>().ToList());
                Cursor.Current = Cursors.WaitCursor;
                ThisWorkbook.VSTATUS = VStatus.CALCULATED;
                var consExporter = new ExporterConsumption();
                consExporter.ExportAtcConsumption(calculatedDDDConsumptions);
                Cursor.Current = Cursors.Default;
                return true;
            }
            else
            {
                MessageBox.Show("There are some issues in the consumption data. Cannot proceed further.\nPlease enter correct data and Validate again.",
                    "No data calculated",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                ThisWorkbook.VSTATUS = VStatus.DIRTY;
                return false;
            }
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
            ThisWorkbook.VSTATUS = VStatus.NA;
        }

        private void Validate_Click(object sender, EventArgs e)
        {
            if (optionValue == 0)
            {
                MessageBox.Show("Please select an appropriate option.");
                return;
            }

            try
            {
                // Show the loading dialog
                CustomLoadingBox.Show("Validation in progress, please wait...");

                Excel.Worksheet dataSheet;
                if (optionValue == OPTION_PRODUCTS)
                {
                    dataSheet = LoadWorkSheet(TemplateFormat.PRODUCT_SHEETNAME);
                } else
                {
                    dataSheet = LoadWorkSheet(TemplateFormat.SUBSTANCE_SHEETNAME);
                }
                //Reset the data sheet..!
                ResetExcelSheet(dataSheet);

                //Done upto this, all the three user data sheets are properly parsed into the objects..!
                Excel.Worksheet availabilityDataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.AVAILABILITY_SHEETNAME];
                Excel.Worksheet structureDataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.HOSPITAL_STRUCTURE_SHEETNAME];
                Excel.Worksheet activityDataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.HOSPITAL_ACTIVITY_SHEETNAME];
                bool force = true;

                parseCheckAvailabilityData = AvailabilityStructureActivityDataParser.ParseAvailability(availabilityDataSheet);
                if (parseCheckAvailabilityData)
                {
                    parseCheckHospitalStructure = AvailabilityStructureActivityDataParser.ParseHospitalStructure(structureDataSheet);
                    if (parseCheckHospitalStructure)
                    {
                        parseCheckActivityData = AvailabilityStructureActivityDataParser.ParseActivity(activityDataSheet);
                        if (!parseCheckActivityData)
                        {
                            checkValidation = false;
                            return;
                        }
                        else
                        {
                            if (ThisWorkbook.AvailabilityData.Count <= 0)
                            {
                                force = false;
                                MessageBox.Show("Hospital Availability Data sheet is empty or the data is not entered properly");
                                checkValidation = false;

                            }
                            else if (ThisWorkbook.HospitalStructureData.Count <= 0 && force)
                            {
                                force = false;
                                MessageBox.Show("Hospital Structure Data sheet is empty or the data is not entered properly");
                                checkValidation = false;
                            }
                            else if (ThisWorkbook.HospitalActivityData.Count <= 0 && force)
                            {
                                force = false;
                                MessageBox.Show("Hospital Activity Data sheet is empty or the data is not entered properly");
                                checkValidation = false;

                            }
                            ParseAndValidateMedicines();
                            checkValidation = true;
                        }
                    }
                    else
                    {
                        checkValidation = false;
                        return;
                    }
                }
                else
                {
                    checkValidation = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                checkValidation = false;
                // Handle exceptions
                MessageBox.Show($"An unexpected error occurred: {ex.Message}");
            }
            finally
            {
                // Close the loading dialog
                CustomLoadingBox.Close();
            }
        }


        private void ParseAndValidateMedicines()
        {
            //Check for the sheets if they are empty or not 
            List<IMedicine> medicineList;
            if (optionValue == 1)
            {
                var dataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.PRODUCT_SHEETNAME];
                ThisWorkbook.ProductData = ProductDataParser.ParseProducts(dataSheet);



                medicineList = ThisWorkbook.ProductData.Cast<IMedicine>().ToList();


            } else
            {
                var dataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.SUBSTANCE_SHEETNAME];
                ThisWorkbook.SubstanceData = SubstanceDataParser.ParseSubstances(dataSheet);
                medicineList = ThisWorkbook.SubstanceData.Cast<IMedicine>().ToList();
            }
            if (medicineList.Count > 0)
            {
                foreach (var med in medicineList)
                {
                    //Call the validateProduct method
                    // pr.ValidateProduct(pr, true); OLD CODE
                    med.Validate(false);

                    if (med.IsValid())
                    {
                        if (med.ATC5.Code != AMUConstants.ATC_Z99_CODE)
                        {
                            if (!Utils.IsATCClassInAvailability(med.ATCClass, ThisWorkbook.AvailabilityData ))
                            {
                                med.AddWarningMsg("The ATC code is not among the antimicrobial classes specified in the availability data. It will be excluded from the calculation/export.");
                                med.SetValidate(true, Medicine.ATCCLASS_VALIDATION);
                            }
                        }
                    }
                    else
                    {
                        if (med.GetValidate(Medicine.ATC5_VALIDATION) && med.ATC5.Code != AMUConstants.ATC_Z99_CODE)
                        {
                            if (!Utils.IsATCClassInAvailability(med.ATCClass, ThisWorkbook.AvailabilityData))
                            {
                                med.AddWarningMsg("The ATC code is not among the antimicrobial classes specified in the availability data. It will be excluded from the calculation/export.");
                                med.SetValidate(true, Medicine.ATCCLASS_VALIDATION);
                            }
                        }
                    }
                }

                // Call the displayProduct method
                DisplayMedicinesAndStatuses(medicineList);
                ////Set isProductValidating to false
                // isMedicineValidating = false;
                Validator.SetStatus(VStatus.PARSED);
            }
            else
            {
                if (optionValue == 1)
                {
                    MessageBox.Show("Product Data sheet is empty.");
                }
                else
                {
                    MessageBox.Show("Substance Data sheet is empty.");
                }
            }
        }
            


        private void DisplayMedicinesAndStatuses(List<IMedicine> listMedicines) // TODO: To be moved out of the Product class
        {
            Excel.Workbook wb = Globals.ThisWorkbook.Application.ActiveWorkbook;
            Excel.Worksheet dataWorksheet;

            int totalRows = listMedicines.Count;
            int totalColumns = TemplateFormat.DATA_SHEET_AUTO_CAL_COL_NB; // Assuming there are 7 columns to display
            int dataStartColumn;
            int dataEndColumn;

            string medicineType;
            string medicinesType;

            IDictionary<string, int> templateColIdxMap;

            if (optionValue == 1)
            {
                dataWorksheet = wb.Worksheets[TemplateFormat.PRODUCT_SHEETNAME];
                dataStartColumn = TemplateFormat.PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX;
                dataEndColumn = TemplateFormat.PRODUCT_DATA_SHEET_AUTO_CALC_END_COL_IDX;
                medicineType = "Product";
                medicinesType = "products";
                templateColIdxMap = TemplateFormat.PRODUCT_COL_IDX_MAP;
            } else
            {
                dataWorksheet = wb.Worksheets[TemplateFormat.SUBSTANCE_SHEETNAME];
                dataStartColumn = TemplateFormat.SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX;
                dataEndColumn = TemplateFormat.SUBSTANCE_DATA_SHEET_AUTO_CALC_END_COL_IDX;
                medicineType = "Substance";
                medicinesType = "substances";
                templateColIdxMap = TemplateFormat.SUBSTANCE_COL_IDX_MAP;
            }
            //Logic:
            //Instead of interacting with the excel object for every product, create a data object and populate that, after populating the obj, assign it to the Excel Range..!
            
            // Create a 2D array for data..!
            //For Rows and Columns..!
            object[,] data = new object[totalRows, totalColumns];
            // Create a 2D array for status data..!
            object[,] statusData = new object[totalRows, 2];
            bool hasErrors = false;

            // Clear the existing formatting and values for the status columns
            Excel.Range statusRange = dataWorksheet.Range[dataWorksheet.Cells[listMedicines[0].LineNo, 1],
                                                             dataWorksheet.Cells[listMedicines[totalRows - 1].LineNo, 2]];
            statusRange.Interior.Color = Color.White; // Reset cell background color to white
            statusRange.Value2 = ""; // Clear existing messages

            for (int i = 0; i < totalRows; i++)
            {
                IMedicine med = listMedicines[i];
                int lineNo = med.LineNo;

                // Populate data array
                if (med.GetValidate(Medicine.CONVERSION_VALIDATION) && med.ConversionFactor != 1)
                {
                    data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_CONVERSION_FACTOR_COL_IDX]-1] = med.ConversionFactor;
                }
                if (med.GetValidate(Medicine.CONTENT_VALIDATION))
                {
                    data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_CONTENT_COL_IDX] - 1] = med.Content.Value;
                    data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_CONTENT_UNIT_COL_IDX] - 1] = med.Content.Unit.Code;
                }
                if (med.GetValidate(Medicine.ARS_VALIDATION))
                {
                    data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_ARS_COL_IDX] - 1] = med.ARS;
                }
                if (med.GetValidate(Medicine.DDD_VALIDATION))
                {
                    data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_DDD_VALUE_COL_IDX] - 1] = med.DDD.Value;
                    data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_DDD_UNIT_COL_IDX] - 1] = med.DDD.Unit.Code;

                }
                if (med.GetValidate(Medicine.CONVERSION_VALIDATION) && med.GetValidate(Product.CONTENT_VALIDATION) && med.GetValidate(Product.DDD_VALIDATION))
                {
                    data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_DPP_COL_IDX] - 1] = med.NbDDD;
                }

                // Safe null checks for pr.AWaRe and pr.MEML
                data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_AWR_COL_IDX] - 1] = med.AWaRe ?? "N/A"; // Default to "N/A" if pr.AWaRe is null
                data[i, templateColIdxMap[TemplateFormat.AUTO_CALC_MEML_COL_IDX] - 1] = YesNoNAString.GetStringFromYesNoNA(med.MEML);

                // Populate statusData array..!
                switch (med.Status)
                {
                    case EntityStatus.OK:
                        statusData[i, TemplateFormat.DATA_SHEET_STATUS_COL_IDX - 1] = "OK";
                        statusData[i, TemplateFormat.DATA_SHEET_STATUS_MSG_COL_IDX - 1] = "";
                        dataWorksheet.Range[dataWorksheet.Cells[lineNo, 1], dataWorksheet.Cells[lineNo, 2]].Interior.Color = Color.Green;
                        dataWorksheet.Range[dataWorksheet.Cells[lineNo, 1], dataWorksheet.Cells[lineNo, 2]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        break;
                    case EntityStatus.WARNING:
                        statusData[i, TemplateFormat.DATA_SHEET_STATUS_COL_IDX - 1] = "WARNING";
                        statusData[i, TemplateFormat.DATA_SHEET_STATUS_MSG_COL_IDX - 1] = med.GetStatusMessages();
                        dataWorksheet.Range[dataWorksheet.Cells[lineNo, 1], dataWorksheet.Cells[lineNo, 2]].Interior.Color = Color.Orange;
                        dataWorksheet.Range[dataWorksheet.Cells[lineNo, 1], dataWorksheet.Cells[lineNo, 2]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        break;
                    //case EntityStatus.INFO:
                    case EntityStatus.ERROR:
                        string statusText = (med.Status == EntityStatus.INFO) ? "INFO" : "ERR";
                        string msgTxt = med.GetStatusMessages();
                        statusData[i, TemplateFormat.DATA_SHEET_STATUS_COL_IDX - 1] = statusText;
                        statusData[i, TemplateFormat.DATA_SHEET_STATUS_MSG_COL_IDX - 1] = msgTxt;
                        dataWorksheet.Range[dataWorksheet.Cells[lineNo, 1], dataWorksheet.Cells[lineNo, 2]].Interior.Color = (med.Status == EntityStatus.ERROR) ? Color.Red : Color.Yellow;
                        dataWorksheet.Range[dataWorksheet.Cells[lineNo, 1], dataWorksheet.Cells[lineNo, 2]].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        foreach (var validation in med.ValidationMessages)
                        {
                            hasErrors = true;
                            dynamic dataField = validation.ErrorField;
                            if (dataField != null)
                            {
                                int columnIndex = dataField.FieldColumn;
                                Excel.Range cellWithIssue = dataWorksheet.Cells[lineNo, columnIndex];
                                cellWithIssue.Borders.Color = Color.OrangeRed;
                                cellWithIssue.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            }
                        }

                        // Set the concatenated validation messages..!
                        dataWorksheet.Cells[lineNo, 2].Value = msgTxt;
                        break;
                }
            }


            // Set the entire range at once for both data and status
            dataWorksheet.Range[dataWorksheet.Cells[listMedicines[0].LineNo, dataStartColumn], dataWorksheet.Cells[listMedicines[totalRows - 1].LineNo, dataEndColumn]].Value2 = data;
            dataWorksheet.Range[dataWorksheet.Cells[listMedicines[0].LineNo, 1], dataWorksheet.Cells[listMedicines[totalRows - 1].LineNo, 2]].Value2 = statusData;

            if (hasErrors)
            {
                // Show products with errors in the MessageBox..!
                MessageBox.Show($"Validation of the {medicinesType} has been completed.\n\nThere are some errors in the {medicineType} data sheet.\n\nPlease review the data and validate it again.");
            }
            else
            {
                MessageBox.Show($"Validation of the {medicinesType} has been successfully completed.\n\nPlease proceed by pressing the Calculate Use button to calculate the use in DDD.\r\n");
            }
            // Autofit columns after the loop
            dataWorksheet.Columns.AutoFit();
            dataWorksheet.Application.ActiveWorkbook.Save();
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet3_Startup);
            this.Shutdown += new System.EventHandler(Sheet3_Shutdown);
        }

        #endregion

    }
}
