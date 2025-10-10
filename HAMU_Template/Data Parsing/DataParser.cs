// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Constants;
using AMU_Template.Models;
using AMU_Template.Parsers;
using AMU_Template.Validations;
using HAMU_Template.Constants;
using HAMU_Template.Helper;
using HAMU_Template.Models;
using HAMU_Template.Models.Mappings;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Excel = Microsoft.Office.Interop.Excel;



namespace HAMU_Template.Data_Parsing
{


    public static class HAMUCommonParser
    {
        public static FacilityStructureLevel? ParseStructureLevel(string val, string variable, ErrorStatus es, ref string variableErrors, bool mandatory, FacilityStructureLevel? defValue=null)
        {
            if (val == null)
            {
                if (mandatory)
                {
                    if (defValue == null)
                    {
                        string errMsg = $"{variable} is mandatory.";
                        variableErrors += errMsg + "\n";
                        es.Status = EntityStatus.ERROR;
                        es.AddErrorMsgs(errMsg);
                        return null;
                    }
                    else
                    {
                        return defValue;
                    }
                }
                return null;
            }

            string sLevelStr = val.ToString().ToUpper().Trim();
            FacilityStructureLevel sLevel = FacilityStructureLevelString.GetFacilityStructureLevelForString(sLevelStr);

            return sLevel;
        }
    }

    public static class ReferenceDataParser
    {
        public static Dictionary<string, DDDCombination> listCombDdds = new Dictionary<string, DDDCombination>();
        public static Dictionary<string, DDD> listDdds = new Dictionary<string, DDD>();
        public static Dictionary<string, MeasureUnit> listUnits = new Dictionary<string, MeasureUnit>();
        public static Dictionary<string, AdministrationRoute> listRoAs = new Dictionary<string, AdministrationRoute>(); // need to confirm with DP
        public static Dictionary<string, ATC> listAtcs = new Dictionary<string, ATC>();
        public static Dictionary<string, Salt> listSalts = new Dictionary<string, Salt>();

        public static List<ATC> ProcessATC(Worksheet workSheet)
        {
            Range usedRange = workSheet.UsedRange;
            return ATCListParser.ParseATCList(usedRange);
        }

        public static List<Aware> ProcessAware(Worksheet workSheet)
        {
            Range usedRange = workSheet.UsedRange;
            return AwareListParser.ParseAwareList(usedRange);
        }

        public static List<MEML> ProcessMeml(Worksheet workSheet)
        {
            Range usedRange = workSheet.UsedRange;
            return MEMLListParser.ParseMEMLList(usedRange);

        }

        public static List<DDDCombination> ProcessDDDCombination(Worksheet workSheet, Dictionary<string, ATC> atc5_dict, Dictionary<string, AdministrationRoute> roa_dict, Dictionary<string, MeasureUnit> unit_dict)
        {
            Range usedRange = workSheet.UsedRange;
            return DDDCombinationListParser.ParseDDDCombinationList(usedRange, atc5_dict, roa_dict, unit_dict);
        }

        public static List<ConversionFactor> ProcessConversionFactor(Worksheet workSheet, Dictionary<string, ATC> atc5_dict, Dictionary<string, AdministrationRoute> roa_dict, Dictionary<string, Salt> salt_dict, Dictionary<string, MeasureUnit> unit_dict)
        {
            Range usedRange = workSheet.UsedRange;
            return ConversionFactorListParser.ParseConversionFactorList(usedRange, atc5_dict, roa_dict, salt_dict, unit_dict);
        }

        public static List<DDD> ProcessDDD(Worksheet workSheet, Dictionary<string, ATC> atc5_dict, Dictionary<string, AdministrationRoute> roa_dict, Dictionary<string, Salt> salt_dict, Dictionary<string, MeasureUnit> unit_dict)
        {
            Range usedRange = workSheet.UsedRange;
            return DDDListParser.ParseDDDList(usedRange, atc5_dict, roa_dict, salt_dict, unit_dict);
        }

        public static List<AdministrationRoute> ProcessRoAs(Worksheet workSheet)
        {
            Range usedRange = workSheet.UsedRange;
            return RoaListParser.ParseRoaList(usedRange);
        }

        public static List<Salt> ProcessSalt(Worksheet workSheet)
        {
            Range usedRange = workSheet.UsedRange;
            return SaltListParser.ParseSaltList(usedRange);
        }

        public static List<MeasureUnit> ProcessUnit(Worksheet workSheet)
        {

            Range usedRange = workSheet.UsedRange;
            return MeasureUnitListParser.ParseMeasureUnitList(usedRange);
        }

        public static List<ProductOrigin> ProcessProductOrigin(Worksheet workSheet)
        {

            Range usedRange = workSheet.UsedRange;
            return ProductOriginListParser.ParseProductOriginList(usedRange);
        }
    }

    public static class AvailabilityStructureActivityDataParser
    {

        //Parsing and Processing the Availability data and storing them in an object of list type <Class - Availability> ..!
        //COUNTRY       HOSPITAL      YEAR      A07AA       D01BA      J01     J02     J04     J05     P01AB   P01B    LEVEL

        public static bool ParseAvailability(Worksheet workSheet)
        {
            List<Availability> availabilityData = new List<Availability>();

            string country;
            string hospital;
            int? year;
            FacilityStructureLevel? level;
            bool? availabilityA07AA;
            bool? availabilityD01BA;
            bool? availabilityJ01;
            bool? availabilityJ02;
            bool? availabilityJ04;
            bool? availabilityJ05;
            bool? availabilityP01AB;
            bool? availabilityP01B;
            List<String> availabilities;


            Excel.Range usedRange = workSheet.UsedRange;

            // Get the actual rows..!
            int nbRows = Utils.GetRowsCountAvailabilityData(usedRange);

            if(nbRows ==0)
            {
                MessageBox.Show($"There is no Availability data defined in the worksheet `{TemplateFormat.AVAILABILITY_SHEETNAME}`");
                return false;
            }

            ErrorStatus errorStatus = new ErrorStatus();
            string variableErrors = "\n";

            errorStatus.Reset();

            for (int row = 2; row <= nbRows + 1; row++)
            {
                country = CommonParser.ParseString(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_COUNTRY_COL_IDX].Value2, "Country", errorStatus, ref variableErrors, true);
                hospital = CommonParser.ParseString(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_HOSPITAL_COL_IDX].Value2, "Hospital", errorStatus, ref variableErrors, true);
                year = CommonParser.ParseYear(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_YEAR_COL_IDX].Value2, "Year", errorStatus, ref variableErrors, true);
                level = HAMUCommonParser.ParseStructureLevel(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_LEVEL_COL_IDX].Value2, "Level", errorStatus, ref variableErrors, true, FacilityStructureLevel.Hospital);
                availabilityA07AA = CommonParser.ParseBoolean(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_ATC_COL_IDX_MAP["A07AA"]].Value2, "AvailabilityA07AA", errorStatus, ref variableErrors, true, false);
                availabilityD01BA = CommonParser.ParseBoolean(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_ATC_COL_IDX_MAP["D01BA"]].Value2, "AvailabilityD01BA", errorStatus, ref variableErrors, true, false);
                availabilityJ01 = CommonParser.ParseBoolean(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_ATC_COL_IDX_MAP["J01"]].Value2, "AvailabilityJ01", errorStatus, ref variableErrors, true, false);
                availabilityJ02 = CommonParser.ParseBoolean(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_ATC_COL_IDX_MAP["J02"]].Value2, "AvailabilityJ02", errorStatus, ref variableErrors, true, false);
                availabilityJ04 = CommonParser.ParseBoolean(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_ATC_COL_IDX_MAP["J04"]].Value2, "AvailabilityJ04", errorStatus, ref variableErrors, true, false);
                availabilityJ05 = CommonParser.ParseBoolean(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_ATC_COL_IDX_MAP["J05"]].Value2, "AvailabilityJ05", errorStatus, ref variableErrors, true, false);
                availabilityP01AB = CommonParser.ParseBoolean(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_ATC_COL_IDX_MAP["P01AB"]].Value2, "AvailabilityP01AB", errorStatus, ref variableErrors, true, false);
                availabilityP01B = CommonParser.ParseBoolean(usedRange.Cells[row, TemplateFormat.AVAILABILITY_SHEET_ATC_COL_IDX_MAP["P01B"]].Value2, "AvailabilityP01B", errorStatus, ref variableErrors, true, false);

                if (errorStatus.Status != EntityStatus.OK)
                {
                    MessageBox.Show($"There are some errors in availability data sheet, {variableErrors}\nPlease check at row number \"{row}\" and validate again");
                    return false;
                }

                availabilities = new List<string>();
                if (availabilityA07AA==true)
                {
                    availabilities.Add("A07AA");
                }
                if (availabilityD01BA == true)
                {
                    availabilities.Add("D01BA");
                }
                if (availabilityJ01== true)
                {
                    availabilities.Add("J01");
                }
                if (availabilityJ02 == true)
                {
                    availabilities.Add("J02");
                }
                if (availabilityJ04 == true)
                {
                    availabilities.Add("J04");
                }
                if (availabilityJ05 == true)
                {
                    availabilities.Add("J05");
                }
                if (availabilityP01AB == true)
                {
                    availabilities.Add("P01AB");
                }
                if (availabilityP01B == true)
                {
                    availabilities.Add("P01B");
                }

                
                foreach (string atcClass in availabilities )
                {

                    // Creating our own custom key..!
                    AvailabilityKey key = new AvailabilityKey(country, (int)year, hospital, FacilityStructureLevel.Hospital, atcClass);
               
                    //Check for redundant enteries if the data is already added then throw an error..!
                    if (availabilityData.Any(item => item.Key == key))
                    {
                        MessageBox.Show($"Availability data already defined for country {country}, hospital {hospital}, year {year} and ATC Class {atcClass}!");
                        availabilityData.Clear();
                        return false;
                    }

                    Availability avail = new Availability
                    {
                        Country = country,
                        Hospital = hospital,
                        Year = (int)year,
                        ATCClass = atcClass,
                        Level = FacilityStructureLevel.Hospital
                    };
                    availabilityData.Add(avail);
                }
                
                errorStatus.Reset();
            }

            ThisWorkbook.AvailabilityData = availabilityData;
            return true;
        }

        public static bool ParseHospitalStructure(Excel.Worksheet workSheet)
        {
            List<HospitalStructure> hospitalStructureData = new List<HospitalStructure>();

            string country;
            string hospital;
            int? year;

            Excel.Range usedRange = workSheet.UsedRange;

            // Get the actual rows..!
            int rowsWithData = Utils.GetRowsCountHospitalStructure(usedRange);

            ErrorStatus errorStatus = new ErrorStatus();
            string variableErrors = "\n";

            errorStatus.Reset();

            for (int row = 2; row <= rowsWithData + 1; row++)
            {
                country = CommonParser.ParseString(usedRange.Cells[row, TemplateFormat.STRUCTURE_SHEET_COUNTRY_COL_IDX].Value2, "Country", errorStatus, ref variableErrors, true);
                hospital = CommonParser.ParseString(usedRange.Cells[row, TemplateFormat.STRUCTURE_SHEET_HOSPITAL_COL_IDX].Value2, "Hospital", errorStatus, ref variableErrors, true);
                year = CommonParser.ParseYear(usedRange.Cells[row, TemplateFormat.STRUCTURE_SHEET_YEAR_COL_IDX].Value2, "Year", errorStatus, ref variableErrors, true);
                
                if (errorStatus.Status != EntityStatus.OK)
                {
                    MessageBox.Show($"There are some errors in Hospital structure data sheet, {variableErrors}\nPlease check at row number \"{row}\" and validate again");
                    hospitalStructureData.Clear();
                    return false;
                }

                // Creating our own custom key..!
                HospitalStructureKey hKey = new HospitalStructureKey(country, (int)year, hospital);

                // Check if the key already exists in the hospitalStructureMaps..!
                if (hospitalStructureData.Any(item => item.Key == hKey))
                {
                    MessageBox.Show($"Hospital structure data already defined for country {country}, hospital {hospital}, and year {year}");
                    hospitalStructureData.Clear();
                    return false;
                }

                hospitalStructureData.Add(new HospitalStructure
                {
                    Country = country,
                    Year = year,
                    Hospital = hospital,
                });
                errorStatus.Reset();
            }
            ThisWorkbook.HospitalStructureData = hospitalStructureData;

            return true;
        }

        public static bool ParseActivity(Excel.Worksheet workSheet)
        {

            List<HospitalActivity> hospitalActivityData = new List<HospitalActivity>();

            string country;
            string hospital;
            int? year;
            string structure;
            FacilityStructureLevel level;
            double? patientDays;
            double? admissions;
            HospitalActivityKey cKey;

            


            Excel.Range usedRange = workSheet.UsedRange;

            // Get the actual rows..!
            int? rowsWithData = Utils.GetRowsCountHospitalActivityData(usedRange);

            ErrorStatus errorStatus = new ErrorStatus();
            string variableErrors = "\n";

            for (int row = 2; row <= rowsWithData + 1; row++)
            {

                country = CommonParser.ParseString(usedRange.Cells[row, TemplateFormat.ACTIVITY_SHEET_COUNTRY_COL_IDX].Value2, "Country", errorStatus, ref variableErrors, true);
                hospital = CommonParser.ParseString(usedRange.Cells[row, TemplateFormat.ACTIVITY_SHEET_HOSPITAL_COL_IDX].Value2, "Hospital", errorStatus, ref variableErrors, true);
                year = CommonParser.ParseYear(usedRange.Cells[row, TemplateFormat.ACTIVITY_SHEET_YEAR_COL_IDX].Value2, "Year", errorStatus, ref variableErrors, true);
                structure = HAMUCommonParser.ParseStructureLevel(usedRange.Cells[row, TemplateFormat.ACTIVITY_SHEET_STRUCTURE_COL_IDX].Value2, "Structure", errorStatus, ref variableErrors, false);
                patientDays = CommonParser.ParseNumber(usedRange.Cells[row, TemplateFormat.ACTIVITY_SHEET_PATIENT_DAYS_COL_IDX].Value2, "PatientDays", errorStatus, ref variableErrors, true);
                admissions = CommonParser.ParseNumber(usedRange.Cells[row, TemplateFormat.ACTIVITY_SHEET_ADMISSIONS_COL_IDX].Value2, "Admissions", errorStatus, ref variableErrors, true);

                

                if (errorStatus.Status != EntityStatus.OK)
                {
                    MessageBox.Show($"There are some errors in Hospital Activity data sheet, {variableErrors}\nPlease check at row number \" {row} \" and validate again");
                    hospitalActivityData.Clear();
                    return false;
                }

                if (!ThisWorkbook.HospitalStructureData.Any(item => item.Country == country && item.Year == year && item.Hospital == hospital))
                {
                    hospitalActivityData.Clear();
                    MessageBox.Show($"Error parsing activity, given hospital {hospital} is not defined for the given year {year}");
                    return false;
                }

                if (!ThisWorkbook.AvailabilityData.Any(item => item.Country == country && item.Year == year && item.Hospital == hospital))
                {
                    hospitalActivityData.Clear();
                    MessageBox.Show($"Error parsing activity, given hospital {hospital} and year {year} don't have availability data.");
                    return false;
                }

                //Multiple structure is yet to be implemented..! Only manage Hospital at the moment

                HospitalActivity activity = new HospitalActivity
                {
                    Country = country,
                    Year = (int)year,
                    Hospital = hospital,
                    Level = FacilityStructureLevel.Hospital,
                    Structure = "__HOSPITAL__",
                    PatientDays = (int)patientDays,
                    Admissions = (int)admissions
                };

                if (hospitalActivityData.Any(item => item.Key == activity.Key))
                {
                    hospitalActivityData.Clear();
                    MessageBox.Show($"Error parsing activity, there are duplication of activity for hospital {hospital} and year {year}.");
                    return false;
                }

                hospitalActivityData.Add(activity);
                errorStatus.Reset();
            }

            ThisWorkbook.HospitalActivityData = hospitalActivityData;

            return true;
        }
    }


    public class MedicineDataParser
    {
        public static DataField<T> InitializeDataField<T>(string varName, bool mandatory, int colIndex)
        {
            return new DataField<T>
            {
                Name = varName,
                IsMandatory = mandatory,
                IsValid = true,
                FieldColumn = colIndex
            };
        }

        private static string ParseCountryIso(string value)
        {
            if (value.Length != 3)
            {
                throw new ArgumentException($"Country ISO3 {value} value is invalid. It must be 3-letters code.");
            }
            return value;
        }

        protected static string ParseNonEmptyString(string value)
        {
            if (String.IsNullOrEmpty(value))
            {
                throw new ArgumentException($"Empty value. It is mandatory.");
            }
            return value;
        }

        protected static Decimal ParseDecimal(string value)
        {
            return Convert.ToDecimal(value);
        }

        protected static AdministrationRoute ParseRoa(string value)
        {
            if (!ThisWorkbook.AdminRouteDataDict.ContainsKey(value))
            {
                throw new ArgumentException($"ROA value {value} is invalid.");
            }
            return ThisWorkbook.AdminRouteDataDict[value];
        }

        protected static MeasureUnit ParseMeasureUnit(string value)
        {
            if (!ThisWorkbook.UnitDataDict.ContainsKey(value))
            {
                throw new ArgumentException($"MeasureUnit value {value} is invalid.");
            }
            return ThisWorkbook.UnitDataDict[value];
        }

        protected static ATC ParseAtc5(string value)
        {
            if (value == AMUConstants.ATC_Z99_CODE)
            {
                return AMUConstants.ATC_Z99;
            }
            if (!ThisWorkbook.ATC5DataDict.ContainsKey(value))
            {
                throw new ArgumentException($"ATC5 value {value} is invalid.");
            }
            return ThisWorkbook.ATC5DataDict[value];
        }

        protected static Salt ParseSalt(string value)
        {
            if (!ThisWorkbook.SaltDataDict.ContainsKey(value))
            {
                throw new ArgumentException($"Salt value {value} is invalid.");
            }
            return ThisWorkbook.SaltDataDict[value];
        }

        protected static DDDCombination ParseCombination(string value)
        {
            if (value == AMUConstants.COMB_Z99_CODE)
            {
                return AMUConstants.COMB_Z99;
            }
            if (!ThisWorkbook.DDDCombinationDataDict.ContainsKey(value))
            {
                throw new ArgumentException($"Combination value {value} is invalid.");
            }
            return ThisWorkbook.DDDCombinationDataDict[value];
        }

        protected static int ParseYear(string value)
        {
            int y = Convert.ToInt32(value);
            if (y < 1970 || y > DateTime.Today.Year)
            {
                throw new ArgumentException($"Year {y} is invalid [1970-{DateTime.Today.Year}]");
            }
            return y;
        }

        public static void ParseMedicineCountryISO3(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Medicine.COUNTRY_FIELD, true, colIdx);

            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value);
                value2 = value2.ToUpper();
                try
                {
                    string co = ParseCountryIso(value2);
                    df.InputValue = value;
                    df.Value = co;
                    df.IsValid = true;
                    med.SetField(Medicine.COUNTRY_FIELD, df);
                    med.Country = co;
                }
                catch (ArgumentException e)
                {
                    errMsg = $"{Medicine.COUNTRY_FIELD}=>{e.Message}";
                    med.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    med.SetField(Medicine.COUNTRY_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.COUNTRY_FIELD} is mandatory.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                med.SetField(Medicine.COUNTRY_FIELD, df);
            }
        }

        public static void ParseMedicineHospital(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Medicine.HOSPITAL_FIELD, true, colIdx);

            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value);
                df.InputValue = value;
                df.Value = value2;
                df.IsValid = true;
                med.SetField(Medicine.HOSPITAL_FIELD, df);

            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.HOSPITAL_FIELD} is mandatory.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                med.SetField(Medicine.HOSPITAL_FIELD, df);
            }
        }

        public static void ParseMedicineLabel(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.LABEL_FIELD, true, colIdx);

            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value);
                df.InputValue = value;
                df.Value = value2;
                df.IsValid = true;
                med.SetField(Medicine.LABEL_FIELD, df);

            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.LABEL_FIELD} is mandatory.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsMissing = true;
                df.IsValid = false;
                med.SetField(Medicine.LABEL_FIELD, df);
            }
        }

        public static void ParseMedicineRoa(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<AdministrationRoute> df = InitializeDataField<AdministrationRoute>(Medicine.ROUTE_ADMIN_FIELD, true, colIdx);

            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value).ToUpper();
                try
                {
                    string value3 = value2.ToUpper();
                    AdministrationRoute roa = ParseRoa(value3);
                    df.InputValue = value;
                    df.Value = roa;
                    df.IsValid = true;
                    med.SetField(Medicine.ROUTE_ADMIN_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the medoduct
                    errMsg = $"{Medicine.ROUTE_ADMIN_FIELD} value is invalid.";
                    med.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    med.SetField(Medicine.ROUTE_ADMIN_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.ROUTE_ADMIN_FIELD} is mandatory.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                med.SetField(Medicine.ROUTE_ADMIN_FIELD, df);
            }
        }

        public static void ParseMedicineStrength(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Decimal> df = InitializeDataField<Decimal>(Medicine.STRENGTH_FIELD, true, colIdx);

            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value);
                try
                {
                    Decimal ps = ParseDecimal(value2);
                    df.InputValue = value;
                    df.Value = ps;
                    df.IsValid = true;
                    med.SetField(Medicine.STRENGTH_FIELD, df);
                }
                catch (FormatException)
                {
                    // Add an error message to the product
                    errMsg = $"{Medicine.STRENGTH_FIELD} value is invalid.";
                    med.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    med.SetField(Medicine.STRENGTH_FIELD, df);
                }

            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.STRENGTH_FIELD} is mandatory.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                med.SetField(Medicine.STRENGTH_FIELD, df);
            }
        }

        public static void ParseMedicineStrengthUnit(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<MeasureUnit> df = InitializeDataField<MeasureUnit>(Medicine.STRENGTH_UNIT_FIELD, true, colIdx);

            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value).ToUpper();
                try
                {
                    string value3 = value2.ToUpper();
                    MeasureUnit mu = ParseMeasureUnit(value3);
                    df.InputValue = value;
                    df.Value = mu;
                    df.IsValid = true;
                    med.SetField(Medicine.STRENGTH_UNIT_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Medicine.STRENGTH_UNIT_FIELD} value is invalid.";
                    med.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    med.SetField(Medicine.STRENGTH_UNIT_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.STRENGTH_UNIT_FIELD} is mandatory.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                med.SetField(Medicine.STRENGTH_UNIT_FIELD, df);
            }
        }

        public static void ParseMedicineConcentrationVolume(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Decimal> df = InitializeDataField<Decimal>(Medicine.CONCENTRATION_VOLUME_FIELD, true, colIdx);

            string errMsg;


            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                med.SetField(Medicine.CONCENTRATION_VOLUME_FIELD, df);
                return;
            }

            try
            {
                Decimal ps = ParseDecimal(value);
                df.InputValue = value;
                df.Value = ps;
                df.IsValid = true;
                med.SetField(Medicine.CONCENTRATION_VOLUME_FIELD, df);
            }
            catch (FormatException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.CONCENTRATION_VOLUME_FIELD} value is invalid.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                med.SetField(Medicine.CONCENTRATION_VOLUME_FIELD, df);
            }
        }

        public static void ParseMedicineVolume(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Decimal> df = InitializeDataField<Decimal>(Medicine.VOLUME_FIELD, true, colIdx);

            string errMsg;


            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                med.SetField(Medicine.VOLUME_FIELD, df);
                return;
            }

            try
            {
                Decimal ps = ParseDecimal(value);
                df.InputValue = value;
                df.Value = ps;
                df.IsValid = true;
                med.SetField(Medicine.VOLUME_FIELD, df);
            }
            catch (FormatException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.VOLUME_FIELD} value is invalid.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                med.SetField(Medicine.VOLUME_FIELD, df);
            }
        }

        public static void ParseMedicineAtc5(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<ATC> df = InitializeDataField<ATC>(Medicine.ATC5_FIELD, true, colIdx);

            string errMsg = "";

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value).ToUpper();
                try
                {
                    string value3 = value2.ToUpper();
                    ATC atc = ParseAtc5(value3);
                    df.InputValue = value;
                    df.Value = atc;
                    df.IsValid = true;
                    med.SetField(Medicine.ATC5_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Medicine.ATC5_FIELD} value is invalid.";
                    med.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    med.SetField(Medicine.ATC5_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.ATC5_FIELD} is mandatory.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                med.SetField(Medicine.ATC5_FIELD, df);
            }
        }

        public static void ParseMedicineSalt(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Salt> df = InitializeDataField<Salt>(Medicine.SALT_FIELD, true, colIdx);

            string errMsg;

            if (string.IsNullOrEmpty(value))
            { // set default salt XXXX
                var defaultSalt = "XXXX";
                Salt salt = ParseSalt(defaultSalt);
                df.InputValue = value;
                df.Value = salt;
                df.IsValid = true;
                df.IsMissing = true;
                med.SetField(Medicine.SALT_FIELD, df);
                return;
            }

            // Check if the value is empty

            string value2 = ParseNonEmptyString(value).ToUpper();
            try
            {
                string value3 = value2.ToUpper();
                Salt salt = ParseSalt(value3);
                df.InputValue = value;
                df.Value = salt;
                df.IsValid = true;
                med.SetField(Medicine.SALT_FIELD, df);
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.SALT_FIELD} value is invalid.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                med.SetField(Medicine.SALT_FIELD, df);
            }
        }

        public static void ParseMedicineCombination(object cellValue, Medicine med, int colIdx)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<DDDCombination> df = InitializeDataField<DDDCombination>(Medicine.COMBINATION_FIELD, true, colIdx);

            string errMsg;

            // Check if the value is empty

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                med.SetField(Medicine.COMBINATION_FIELD, df);
                return;
            }

            try
            {
                string value2 = value.ToUpper();
                DDDCombination comb = ParseCombination(value2);
                df.InputValue = value;
                df.Value = comb;
                df.IsValid = true;
                med.SetField(Medicine.COMBINATION_FIELD, df);
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.COMBINATION_FIELD} value is invalid.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                med.SetField(Medicine.COMBINATION_FIELD, df);
            }
        }

        public static void ParseMedicinePaediatrics(object cellValue, Medicine med, int colIdx)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<YesNoUnknown> df = InitializeDataField<YesNoUnknown>(Medicine.PAEDIATRIC_FIELD, true, colIdx);

            string errMsg;

            try
            {
                string value2 = ParseNonEmptyString(value);
                try
                {
                    string value3 = value2.ToUpper();
                    YesNoUnknown ynu = YesNoUnknownString.GetYesNoUnkFromString(value3);
                    df.InputValue = value;
                    df.Value = ynu;
                    df.IsValid = true;
                    med.SetField(Medicine.PAEDIATRIC_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Medicine.PAEDIATRIC_FIELD} value is invalid.";
                    med.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    med.SetField(Medicine.PAEDIATRIC_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Medicine.PAEDIATRIC_FIELD} is mandatory.";
                med.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                med.SetField(Medicine.PAEDIATRIC_FIELD, df);
            }
        }

        public static void ParseMedicineForm(object cellValue, Medicine med, int colIdx)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Medicine.FORM_FIELD, true, colIdx);

            string errMsg;

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                med.SetField(Medicine.FORM_FIELD, df);
            }
            else
            {
                df.InputValue = value;
                df.Value = value;
                df.IsValid = true;
                df.IsMissing = false;
                med.SetField(Medicine.FORM_FIELD, df);
            }
        }

        public static void ParseMedicineIngredients(object cellValue, Medicine med, int colIdx)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.INGREDIENTS_FIELD, true, colIdx);

            string errMsg;

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                med.SetField(Medicine.INGREDIENTS_FIELD, df);
            }
            else
            {
                df.InputValue = value;
                df.Value = value;
                df.IsValid = true;
                df.IsMissing = false;
                med.SetField(Medicine.INGREDIENTS_FIELD, df);
                if (value.Contains("+") || value.Contains("/"))
                {
                    var infMsg = $"{Medicine.INGREDIENTS_FIELD}=>if you have a fixed dose combination, please use comma to list each INN.";
                    med.AddInfoMsg(infMsg);
                }
            }
        }

        public static Decimal ExtractConsumptionValue(object[,] values, int row, int col, ErrorStatus es, string label, int lineNo, int lastColumn)
        {
            if (col > lastColumn) // Prevent out-of-range error
            {
                es.AddErrorMsgs($"Column index {col} for {label} is out of range.");
                return 0;
            }

            object value = values[row + 1, col]; // Adjust for 1-based Excel index
            if (Decimal.TryParse(Convert.ToString(value), out Decimal result))
            {
                return result;
            }

            if (string.IsNullOrEmpty(Convert.ToString(value)))
            {
                return Decimal.Zero; // Empty cells are treated as zero
            }

            es.AddErrorMsgs($"Value for {label} at line {lineNo} is not valid!");
            return Decimal.Zero;
        }
    }

    public class ProductDataParser: MedicineDataParser
    {

        public static int[] ParseConsYears(Worksheet sheet)
        {
            List<int> years = new List<int>();

            // Determine the last column in the first row
            Range firstRow = sheet.Rows[1] as Range;
            int lColumn = firstRow.Cells[1, sheet.Columns.Count].End[XlDirection.xlToLeft].Column;

            int nbYears = lColumn - TemplateFormat.PRODUCT_DATA_SHEET_CONS_START_COL_IDX + 1;

            // Read years from the relevant columns
            for (int y = 1; y <= nbYears; y++)
            {
                int year = Convert.ToInt32(
                    sheet.Cells[1, TemplateFormat.PRODUCT_DATA_SHEET_CONS_START_COL_IDX + y].Value
                );
                years.Add(year);
            }

            return years.ToArray();
        }

        public static List<ProductConsumption> ParseCons(
        int[] years,
        Worksheet dataSheet,
        List<Product> productData,
        List<Availability> availabilityData,
        List<HospitalActivity> hospitalActivityData,
        List<HospitalStructure> hospitalStructureData,
        ErrorStatus es)
        {

            // Call ParsePackData and return the result
            return ParsePackData(dataSheet, years, productData, availabilityData, hospitalActivityData, hospitalStructureData, es);
        }

        private static List<ProductConsumption> ParsePackData(Worksheet sheet, int[] years,
        List<Product> productData, List<Availability> availabilityData, List<HospitalActivity> activityData,
        List<HospitalStructure> structureData, ErrorStatus es)
        {
            int startRow = TemplateFormat.PRODUCT_DATA_SHEET_START_DATA_ROW_IDX;
            int noOfProducts = Helper.Utils.GetRowsCountProductData(sheet.UsedRange);
            int maxRow = startRow + noOfProducts - 1; // Last row index

            Range yearsColRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 1]];
            Range lastCol = yearsColRange.End[XlDirection.xlToRight];
            int lastColumn = lastCol.Column;

            // Bulk Read: Load entire data range at once
            Range dataRange = sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[maxRow, lastColumn]];

            object[,] values = (object[,])dataRange.Value2;

            var productConsumptionData = new List<ProductConsumption>();

            // 🚀 Process rows in memory
            for (int i = 0; i < noOfProducts; i++)
            {
                int excelRow = startRow + i; // Adjust for Excel row numbering
                ParsePackConsFromArray(values, years, i, excelRow, productData, productConsumptionData, availabilityData, activityData, structureData, es, lastColumn);
            }

            return productConsumptionData;
        }

        private static void ParsePackConsFromArray(object[,] values, int[] years, int index, int excelRow,
        List<Product> products, List<ProductConsumption> productConsumptionData,
        List<Availability> availabilityData, List<HospitalActivity> activityData,
        List<HospitalStructure> structureData, ErrorStatus es, int lastColumn)
        {
            int yCol = 0;
            foreach (int year in years)
            {
                yCol++;
                List<Availability> availData = availabilityData
                        .Where(a => a.Year == year)
                        .ToList();

                if (availData.Count == 0)
                {
                    continue;
                }

                // products = SharedData.Products;

                if (products.Any(p => p.LineNo == excelRow))
                {
                    var pr = products.FirstOrDefault(p => p.LineNo == excelRow);
                    int lineNo = pr.LineNo;
                    int seqNo = pr.SequenceNo;

                    //// Check if ATC5 is null or empty and skip this product if so
                    //if (string.IsNullOrEmpty(pr.ATC5))
                    //{
                    //    es.AddErrorMsgs($"Line {lineNo}: ATC5 code is null or empty. Skipping calculation of this line.");
                    //    continue;
                    //}
                    // check product is valid and DPP is not None
                    if (!pr.IsValid() || pr.NbDDD == Decimal.Zero)
                    {
                        es.AddErrorMsgs($"Line {lineNo}: Product is not valid or does not have PDD. Skipping calculation of this line.");
                        continue;
                    }

                    //var prodCons = productConsumptionData.FirstOrDefault(pc => pc.Year == year && pc.LineNo == lineNo)
                    //               ?? new ProductConsumption();

                    var prodCons = new ProductConsumption();
                    prodCons.LineNo = lineNo;
                    prodCons.Sequence = seqNo;
                    prodCons.ProductId = pr.ProductId;
                    prodCons.UniqueId = pr.UniqueId;
                    prodCons.Label = pr.Label;
                    prodCons.Country = pr.Country;
                    prodCons.Hospital = pr.Hospital;
                    prodCons.Year = year;

                    // At the moment, only managed 
                    prodCons.Level = FacilityStructureLevel.Hospital;
                    prodCons.Structure = "__HOSPITAL__";


                    prodCons.ATC5 = pr.ATC5.Code;
                    prodCons.AMClass = pr.AMClass;
                    prodCons.AtcClass = pr.ATCClass;


                    prodCons.AWaRe = pr.AWaRe;
                    prodCons.MEML = pr.MEML;
                    prodCons.Paediatric = pr.Paediatric;

                    prodCons.Roa = pr.Roa.Code;
                    if (prodCons.Roa == "IS" || prodCons.Roa == "IP")
                    {
                        prodCons.Roa = "I";
                    }

                    prodCons.DPP = pr.NbDDD;

                    var avData = availData.FirstOrDefault(a => a.Country == pr.Country && a.Year == year && a.ATCClass == pr.ATCClass && a.Hospital == pr.Hospital);

                    if (avData != null)
                    {
                        // 🚀 Extract values from `values[,]` instead of accessing Excel cells
                        int baseColIndex = TemplateFormat.PRODUCT_DATA_SHEET_CONS_START_COL_IDX + yCol;

                        prodCons.Packages = ExtractConsumptionValue(values, index, baseColIndex, es, $"PACKAGE {year}", lineNo, lastColumn);
                        prodCons.CalculateDDD();

                        var activity = activityData.FirstOrDefault(a => a.Country == pr.Country && a.Year == year && a.Hospital == pr.Hospital);

                        if (activity != null)
                        {
                            prodCons.Admissions = activity.Admissions;
                            prodCons.BedDays = activity.PatientDays;

                            prodCons.CalculateDDDPerActivity();
                        }
                    }
                    productConsumptionData.Add(prodCons);
                }
            }
        }

        public static List<Product> ParseProducts(Worksheet dataSheet)
        {
            List<Product> listProducts = new List<Product>();

            if (dataSheet == null) return listProducts; // Exit if sheet is missing

            int startRow = TemplateFormat.PRODUCT_DATA_SHEET_START_DATA_ROW_IDX; // First data row
            int noOfProducts = Utils.GetRowsCountProductData(dataSheet.UsedRange); // Get row count
            int maxRow = startRow + noOfProducts - 1; // Last row

            //  Dynamically determine last used column to avoid index out of bounds
            int lastColumn = dataSheet.UsedRange.Columns.Count;
            //  Bulk Read: Load all product data into memory dynamically
            Range dataRange = dataSheet.Range[dataSheet.Cells[startRow, 1], dataSheet.Cells[maxRow, lastColumn]];
            object[,] values = (object[,])dataRange.Value2;

            int i = 0; // Index for `values[,]`
            int j = 1; // Sequence number

            //  Using `while` to match previous row-by-row behavior
            while (i < noOfProducts)
            {
                int excelRow = startRow + i; // Adjust to match Excel row number
                Product pr = ParseProductFromArray(values, i, excelRow); // Parse row

                pr.SequenceNo = j++; // Assign sequence
                listProducts.Add(pr);

                i++; // Move to next row

                // 🚀 Allow UI updates every 100 rows
                if (i % 100 == 0)
                {
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            return listProducts;
        }


        //  Optimized Parsing: Process Data from Array Instead of Excel Cells
        private static Product ParseProductFromArray(object[,] values, int index, int excelRow)
        {
            //  Step 1: Initialize the product object first
            Product pr = new Product
            {
                LineNo = excelRow
            };

            try
            {
                ParseMedicineCountryISO3(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_COUNTRY_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_COUNTRY_COL_IDX);
                ParseMedicineHospital(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_HOSPITAL_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_HOSPITAL_COL_IDX);
                ParseProductId(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_PRODUCT_ID_COL_IDX]), pr);
                ParseMedicineLabel(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_LABEL_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_LABEL_COL_IDX);
                ParseProductPackSize(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_PACKSIZE_COL_IDX]), pr);
                ParseMedicineRoa(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_ROA_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_ROA_COL_IDX);
                ParseMedicineStrength(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_STRENGTH_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_STRENGTH_COL_IDX);
                ParseMedicineStrengthUnit(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_STRENGTH_UNIT_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_STRENGTH_UNIT_COL_IDX);
                ParseMedicineConcentrationVolume(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_CONCENTRATION_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_CONCENTRATION_COL_IDX);
                ParseMedicineVolume(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_VOLUME_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_VOLUME_COL_IDX);
                ParseMedicineAtc5(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_ATC5_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_ATC5_COL_IDX);
                ParseMedicineSalt(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_SALT_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_SALT_COL_IDX);
                ParseMedicineCombination(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_COMBINATION_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_COMBINATION_COL_IDX);
                ParseMedicinePaediatrics(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_PAEDIATRIC_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_PAEDIATRIC_COL_IDX);
                ParseMedicineForm(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_FORM_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_FORM_COL_IDX);
                ParseProductOrigin(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_ORIGIN_COL_IDX]), pr);
                ParseProductGenerics(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_GENERICS_COL_IDX]), pr);
                ParseProductName(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_PRODUCT_NAME_COL_IDX]), pr);
                ParseMedicineIngredients(Convert.ToString(values[index + 1, TemplateFormat.PRODUCT_DATA_SHEET_INGREDIENTS_COL_IDX]), pr, TemplateFormat.PRODUCT_DATA_SHEET_INGREDIENTS_COL_IDX);
            }
            catch (Exception e)
            {
                MessageBox.Show($"An unexpected error occurred when parsing product at row {pr.LineNo}: {e.Message}");
            }
            return pr;
        }

        public static void ParseProductId(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.PRODUCT_ID_FIELD, true, TemplateFormat.PRODUCT_DATA_SHEET_PRODUCT_ID_COL_IDX);
            
            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value);
                df.InputValue = value;
                df.Value = value2;
                df.IsValid = true;
                pr.SetField(Product.PRODUCT_ID_FIELD, df);

            }
            catch(ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.PRODUCT_ID_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                pr.SetField(Product.PRODUCT_ID_FIELD, df);
            }
        }

        public static void ParseProductPackSize(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Decimal> df = InitializeDataField<Decimal>(Product.PACKSIZE_FIELD, true, TemplateFormat.PRODUCT_DATA_SHEET_PACKSIZE_COL_IDX);

            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value);
                try
                {
                    Decimal ps = ParseDecimal(value2);
                    df.InputValue = value;
                    df.Value = ps;
                    df.IsValid = true;
                    pr.SetField(Product.PACKSIZE_FIELD, df);
                }
                catch (FormatException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.PACKSIZE_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.PACKSIZE_FIELD, df);
                }

            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.PACKSIZE_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                pr.SetField(Product.PACKSIZE_FIELD, df);
            }
        }

        public static void ParseProductOrigin(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.PRODUCT_ORIGIN_FIELD, true, TemplateFormat.PRODUCT_DATA_SHEET_ORIGIN_COL_IDX);

            string errMsg;

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.PRODUCT_ORIGIN_FIELD, df);
            }
            else
            {
                string value2 = value.ToUpper();
                if (ThisWorkbook.ProductOriginDataDict.ContainsKey(value2))
                {
                    df.InputValue = value;
                    df.Value = value2;
                    df.IsValid = true;
                    df.IsMissing = false;
                    pr.SetField(Product.PRODUCT_ORIGIN_FIELD, df);
                }
                else
                {
                    errMsg = $"{Product.PRODUCT_ORIGIN_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    df.IsMissing = false;
                    pr.SetField(Product.PRODUCT_ORIGIN_FIELD, df);
                }
            }
        }

        public static void ParseProductGenerics(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<YesNoUnknown> df = InitializeDataField<YesNoUnknown>(Product.GENERICS_FIELD, true, TemplateFormat.PRODUCT_DATA_SHEET_GENERICS_COL_IDX);

            string errMsg;

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsMissing = true;
                df.IsValid = true;
                pr.SetField(Product.GENERICS_FIELD, df);
            }
            else
            {
                string value2 =value.ToUpper();
                try
                {
                    YesNoUnknown ynu = YesNoUnknownString.GetYesNoUnkFromString(value2);
                    df.InputValue = value;
                    df.Value = ynu;
                    df.IsValid = true;
                    pr.SetField(Product.GENERICS_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.GENERICS_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.GENERICS_FIELD, df);
                }
            }
        }

        public static void ParseProductName(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.PRODUCT_NAME_FIELD, true, TemplateFormat.PRODUCT_DATA_SHEET_PRODUCT_NAME_COL_IDX);

            string errMsg;

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.PRODUCT_NAME_FIELD, df);
            }
            else
            {
                df.InputValue = value;
                df.Value = value;
                df.IsValid = true;
                df.IsMissing = false;
                pr.SetField(Product.PRODUCT_NAME_FIELD, df);
            }
        }

        
    }

    public class SubstanceDataParser: MedicineDataParser
    {

        public static int[] ParseConsYears(Worksheet sheet)
        {
            List<int> years = new List<int>();

            // Determine the last column in the first row
            Range firstRow = sheet.Rows[1] as Range;
            int lColumn = firstRow.Cells[1, sheet.Columns.Count].End[XlDirection.xlToLeft].Column;

            int nbYears = lColumn - TemplateFormat.SUBSTANCE_DATA_SHEET_CONS_START_COL_IDX + 1;

            // Read years from the relevant columns
            for (int y = 1; y <= nbYears; y++)
            {
                int year = Convert.ToInt32(
                    sheet.Cells[1, TemplateFormat.SUBSTANCE_DATA_SHEET_CONS_START_COL_IDX + y].Value
                );
                years.Add(year);
            }

            return years.ToArray();
        }

        public static List<SubstanceConsumption> ParseCons(
        int[] years,
        Worksheet dataSheet,
        List<Substance> substanceData,
        List<Availability> availabilityData,
        List<HospitalActivity> hospitalActivityData,
        List<HospitalStructure> hospitalStructureData,
        ErrorStatus es)
        {

            // Call ParsePackData and return the result
            return ParseUnitData(dataSheet, years, substanceData, availabilityData, hospitalActivityData, hospitalStructureData, es);
        }

        private static List<SubstanceConsumption> ParseUnitData(Worksheet sheet, int[] years,
        List<Substance> substances, List<Availability> availabilityData, List<HospitalActivity> activityData,
        List<HospitalStructure> structureData, ErrorStatus es)
        {
            int startRow = TemplateFormat.SUBSTANCE_DATA_SHEET_START_DATA_ROW_IDX;
            int noOfSubstances = Helper.Utils.GetRowsCountSubstanceData(sheet.UsedRange);
            int maxRow = startRow + noOfSubstances - 1; // Last row index

            Range yearsColRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 1]];
            Range lastCol = yearsColRange.End[XlDirection.xlToRight];
            int lastColumn = lastCol.Column;

            // Bulk Read: Load entire data range at once
            Range dataRange = sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[maxRow, lastColumn]];

            object[,] values = (object[,])dataRange.Value2;

            var substanceConsumptionData = new List<SubstanceConsumption>();

            // 🚀 Process rows in memory
            for (int i = 0; i < noOfSubstances; i++)
            {
                int excelRow = startRow + i; // Adjust for Excel row numbering
                ParseUnitConsFromArray(values, years, i, excelRow, substances, substanceConsumptionData, availabilityData, activityData, structureData, es, lastColumn);
            }

            return substanceConsumptionData;
        }

        private static void ParseUnitConsFromArray(object[,] values, int[] years, int index, int excelRow,
        List<Substance> substances, List<SubstanceConsumption> substanceConsumptionData,
        List<Availability> availabilityData, List<HospitalActivity> activityData,
        List<HospitalStructure> structureData, ErrorStatus es, int lastColumn)
        {
            int yCol = 0;
            foreach (int year in years)
            {
                yCol++;
                List<Availability> availData = availabilityData
                        .Where(a => a.Year == year)
                        .ToList();

                if (availData.Count == 0)
                {
                    continue;
                }

                // products = SharedData.Products;

                if (substances.Any(s => s.LineNo == excelRow))
                {
                    var subst = substances.FirstOrDefault(p => p.LineNo == excelRow);
                    int lineNo = subst.LineNo;
                    int seqNo = subst.SequenceNo;

                    //// Check if ATC5 is null or empty and skip this product if so
                    //if (string.IsNullOrEmpty(pr.ATC5))
                    //{
                    //    es.AddErrorMsgs($"Line {lineNo}: ATC5 code is null or empty. Skipping calculation of this line.");
                    //    continue;
                    //}
                    // check product is valid and DPP is not None
                    if (!subst.IsValid() || subst.NbDDD == Decimal.Zero)
                    {
                        es.AddErrorMsgs($"Line {lineNo}: Substance is not valid or does not have PDD. Skipping calculation of this line.");
                        continue;
                    }

                    //var prodCons = productConsumptionData.FirstOrDefault(pc => pc.Year == year && pc.LineNo == lineNo)
                    //               ?? new ProductConsumption();

                    var substCons = new SubstanceConsumption();
                    substCons.LineNo = lineNo;
                    substCons.Sequence = seqNo;
                    substCons.Label = subst.Label;
                    substCons.UniqueId = subst.UniqueId;

                    substCons.Country = subst.Country;
                    substCons.Year = year;

                    // At the moment, only managed 
                    substCons.Level = FacilityStructureLevel.Hospital;
                    substCons.Structure = "__HOSPITAL__";


                    substCons.ATC5 = subst.ATC5.Code;
                    substCons.AMClass = subst.AMClass;
                    substCons.AtcClass = subst.ATCClass;


                    substCons.AWaRe = subst.AWaRe;
                    substCons.MEML = subst.MEML;
                    substCons.Paediatric = subst.Paediatric;

                    substCons.Roa = subst.Roa.Code;
                    if (substCons.Roa == "IS" || substCons.Roa == "IP")
                    {
                        substCons.Roa = "I";
                    }

                    substCons.DPP = subst.NbDDD;

                    var avData = availData.FirstOrDefault(a => a.Country == subst.Country && a.Year == year && a.ATCClass == subst.ATCClass && a.Hospital == subst.Hospital);

                    if (avData != null)
                    {
                        // 🚀 Extract values from `values[,]` instead of accessing Excel cells
                        int baseColIndex = TemplateFormat.PRODUCT_DATA_SHEET_CONS_START_COL_IDX + yCol;

                        substCons.Units = ExtractConsumptionValue(values, index, baseColIndex, es, $"UNITS {year}", lineNo, lastColumn);
                        substCons.CalculateDDD();

                        var activity = activityData.FirstOrDefault(a => a.Country == subst.Country && a.Year == year && a.Hospital == subst.Hospital);

                        if (activity != null)
                        {
                            substCons.Admissions = activity.Admissions;
                            substCons.BedDays = activity.PatientDays;

                            substCons.CalculateDDDPerActivity();
                        }
                    }
                    substanceConsumptionData.Add(substCons);
                }
            }
        }
        public static List<Substance> ParseSubstances(Worksheet dataSheet)
        {
            List<Substance> listSubstances = new List<Substance>();

            if (dataSheet == null) return listSubstances; // Exit if sheet is missing

            int startRow = TemplateFormat.SUBSTANCE_DATA_SHEET_START_DATA_ROW_IDX; // First data row
            int noOfSubstances = Utils.GetRowsCountSubstanceData(dataSheet.UsedRange); // Get row count
            int maxRow = startRow + noOfSubstances - 1; // Last row

            //  Dynamically determine last used column to avoid index out of bounds
            int lastColumn = dataSheet.UsedRange.Columns.Count;
            //  Bulk Read: Load all product data into memory dynamically
            Range dataRange = dataSheet.Range[dataSheet.Cells[startRow, 1], dataSheet.Cells[maxRow, lastColumn]];
            object[,] values = (object[,])dataRange.Value2;

            int i = 0; // Index for `values[,]`
            int j = 1; // Sequence number

            //  Using `while` to match previous row-by-row behavior
            while (i < noOfSubstances)
            {
                int excelRow = startRow + i; // Adjust to match Excel row number
                Substance subst = ParseSubstanceFromArray(values, i, excelRow); // Parse row

                subst.SequenceNo = j++; // Assign sequence
                listSubstances.Add(subst);

                i++; // Move to next row

                // 🚀 Allow UI updates every 100 rows
                if (i % 100 == 0)
                {
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            return listSubstances;
        }


        //  Optimized Parsing: Process Data from Array Instead of Excel Cells
        private static Substance ParseSubstanceFromArray(object[,] values, int index, int excelRow)
        {
            //  Step 1: Initialize the product object first
            Substance subst = new Substance()
            {
                LineNo = excelRow
            };

            try
            {
                ParseMedicineCountryISO3(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_COUNTRY_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_COUNTRY_COL_IDX);
                ParseMedicineHospital(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_HOSPITAL_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_HOSPITAL_COL_IDX);
                ParseMedicineLabel(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_LABEL_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_LABEL_COL_IDX);
                ParseMedicineRoa(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_ROA_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_ROA_COL_IDX);
                ParseMedicineStrength(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_STRENGTH_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_STRENGTH_COL_IDX);
                ParseMedicineStrengthUnit(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_STRENGTH_UNIT_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_STRENGTH_UNIT_COL_IDX);
                ParseMedicineConcentrationVolume(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_CONCENTRATION_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_CONCENTRATION_COL_IDX);
                ParseMedicineVolume(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_VOLUME_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_VOLUME_COL_IDX);
                ParseMedicineAtc5(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_ATC5_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_ATC5_COL_IDX);
                ParseMedicineSalt(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_SALT_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_SALT_COL_IDX);
                ParseMedicineCombination(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_COMBINATION_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_COMBINATION_COL_IDX);
                ParseMedicinePaediatrics(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_PAEDIATRIC_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_PAEDIATRIC_COL_IDX);
                ParseMedicineForm(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_FORM_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_FORM_COL_IDX);
                ParseMedicineIngredients(Convert.ToString(values[index + 1, TemplateFormat.SUBSTANCE_DATA_SHEET_INGREDIENTS_COL_IDX]), subst, TemplateFormat.SUBSTANCE_DATA_SHEET_INGREDIENTS_COL_IDX);
            }
            catch (Exception e)
            {
                MessageBox.Show($"An unexpected error occurred when parsing substance at row {subst.LineNo}: {e.Message}");
            }
            return subst;
        }
    }
}
