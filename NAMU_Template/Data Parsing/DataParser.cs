// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using AMU_Template.Constants;
using AMU_Template.Models;
using AMU_Template.Parsers;
using NAMU_Template.Data_Parsing;
using NAMU_Template.Models;
using NAMU_Template.Constants;
using NAMU_Template.Helper;
using EntityStatus = AMU_Template.Validations.EntityStatus;
using AMU_Template.Validations;


namespace NAMU_Template.Data_Parsing
{

    public static class ReferenceDataParser
    {
        public static Dictionary<string, DDDCombination> listCombDdds = new Dictionary<string, DDDCombination>();
        public static Dictionary<string, DDD> listDdds = new Dictionary<string, DDD>();
        public static Dictionary<string, MeasureUnit> listUnits = new Dictionary<string, MeasureUnit>();
        public static Dictionary<string, AdministrationRoute> listRoAs = new Dictionary<string, AdministrationRoute>(); // need to confirm with DP
        public static Dictionary<string, ATC> listAtcs = new Dictionary<string, ATC>();
        public static Dictionary<string, Salt> listSalts = new Dictionary<string, Salt>();

        public static List<ATC> ProcessATC(Microsoft.Office.Interop.Excel.Worksheet workSheet)
        {
            Range usedRange = workSheet.UsedRange;
            return ATCListParser.ParseATCList(usedRange);
        }

        public static List<Aware> ProcessAware(Microsoft.Office.Interop.Excel.Worksheet workSheet)
        {
            Range usedRange = workSheet.UsedRange;
            return AwareListParser.ParseAwareList(usedRange);
        }

        public static List<MEML> ProcessMeml(Microsoft.Office.Interop.Excel.Worksheet workSheet)
        {
            Range usedRange = workSheet.UsedRange;
            return MEMLListParser.ParseMEMLList(usedRange);

        }

        public static List<DDDCombination> ProcessDDDCombination(Microsoft.Office.Interop.Excel.Worksheet workSheet, Dictionary<string, ATC> atc5_dict, Dictionary<string, AdministrationRoute> roa_dict, Dictionary<string, MeasureUnit> unit_dict)
        {
            Range usedRange = workSheet.UsedRange;
            return DDDCombinationListParser.ParseDDDCombinationList(usedRange, atc5_dict, roa_dict, unit_dict);
        }

        public static List<ConversionFactor> ProcessConversionFactor(Microsoft.Office.Interop.Excel.Worksheet workSheet, Dictionary<string, ATC> atc5_dict, Dictionary<string, AdministrationRoute> roa_dict, Dictionary<string, Salt> salt_dict, Dictionary<string, MeasureUnit> unit_dict)
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

        public static List<Country> ProcessCountry(Worksheet workSheet)
        {

            Range usedRange = workSheet.UsedRange;
            return CountryListParser.ParseCountryList(usedRange);
        }

        public static List<ProductOrigin> ProcessProductOrigin(Worksheet workSheet)
        {

            Range usedRange = workSheet.UsedRange;
            return ProductOriginListParser.ParseProductOriginList(usedRange);
        }
    }

    public static class AvailabilityPopulationDataParser
    {
        public static bool ParsePopulation(int maxRowsForPopulation, int firstRow, Range rows,
                                            out List<Population> popYears, List<DataAvailability> avData)
        {
            popYears = new List<Population>();
            string ctry, atcClass;
            HealthSector? sector;
            int? year;
            Decimal? total;
            Decimal? community;
            Decimal? hospital;
            List<Population> pops = new List<Population>();
            bool noPop = true;
            Decimal? cvpop = Decimal.Zero;
            Decimal? hvpop = Decimal.Zero;
            // ATC class array


            ErrorStatus errorStatus = new ErrorStatus();
            string variableErrors = "\n";

            for (int r = firstRow; r <= maxRowsForPopulation; r++)
            {
                errorStatus.Reset();

                // Parse required fields
                ctry = CommonParser.ParseCountryISO3(rows.Cells[r, TemplateFormat.POP_SHEET_COUNTRY_COL_IDX]?.Value, "Country", errorStatus, ref variableErrors, true);
                year = CommonParser.ParseYear(rows.Cells[r, TemplateFormat.POP_SHEET_YEAR_COL_IDX]?.Value, "Year", errorStatus, ref variableErrors, true);
                atcClass = CommonParser.ParseATCClass(rows.Cells[r, TemplateFormat.POP_SHEET_ATCCLASS_COL_IDX]?.Value, "ATC Class", errorStatus, ref variableErrors, false);
                sector = NAMUCommonParser.ParseHealthSector(rows.Cells[r, TemplateFormat.POP_SHEET_SECTOR_COL_IDX]?.Value, "Sector", errorStatus, ref variableErrors, false);
                total = CommonParser.ParseDecimal(rows.Cells[r, TemplateFormat.POP_SHEET_POP_TOTAL_COL_IDX]?.Value, "Total", errorStatus, ref variableErrors, false);
                community = CommonParser.ParseDecimal(rows.Cells[r, TemplateFormat.POP_SHEET_POP_COMMUNITY_COL_IDX]?.Value, "Community", errorStatus, ref variableErrors, false);
                hospital = CommonParser.ParseDecimal(rows.Cells[r, TemplateFormat.POP_SHEET_POP_HOSPITAL_COL_IDX]?.Value, "Hospital", errorStatus, ref variableErrors, false);


                if (String.IsNullOrEmpty(ctry) || year==null || (total==null && community==null && hospital==null))
                {
                    {
                        MessageBox.Show($"In {TemplateFormat.POPULATION_SHEETNAME} worksheet: Missing information for Population at row {r}.");
                        return false;
                    }
                }

                // Handle cases where either `atcClass` or `sector` or both are missing
                if (string.IsNullOrEmpty(atcClass) && sector == null)
                {
                    // Special case where one figure is given for all atcClass and all sector
                    foreach (string atcClassItem in AMUConstants.ATCClasses)
                    {
                        atcClass = atcClassItem;

                        bool av = Utils.IsAmClassYearAvailability(atcClass, (int)year, avData);
                        if (!av) continue;

                        var amClassAvail = Utils.GetATCClassYearAvailability(atcClass, (int)year, avData);
                        var sectors = amClassAvail.Keys.ToList();

                        foreach (var sect in sectors)
                        {
                            var pop = new Population();
                            var sectAvail = amClassAvail[sect];

                            if (sectAvail.AvailabilityTotal)
                            {
                                pop.Country = ctry;
                                pop.Year = (int)year;
                                pop.Sector = sect;
                                pop.ATCClass = atcClass;
                                pop.TotalPopulation = total ?? Decimal.Zero;
                                pop.HospitalPopulation = Decimal.Zero;
                                pop.CommunityPopulation = Decimal.Zero;
                            }
                            else if (sectAvail.AvailabilityCommunity && sectAvail.AvailabilityHospital)
                            {
                                cvpop = (community == Decimal.Zero) ? total : community;
                                hvpop = (hospital == Decimal.Zero) ? total : hospital;
                                pop.Country = ctry;
                                pop.Year = (int)year;
                                pop.ATCClass = atcClass;
                                pop.Sector = sect;
                                pop.TotalPopulation = Decimal.Zero;
                                pop.CommunityPopulation = cvpop;
                                pop.HospitalPopulation = hvpop;
                            }
                            else if (sectAvail.AvailabilityCommunity)
                            {
                                cvpop = (community == Decimal.Zero) ? total : community;
                                pop.Country = ctry;
                                pop.Year = (int)year;
                                pop.ATCClass = atcClass;
                                pop.Sector = sect;
                                pop.TotalPopulation = Decimal.Zero;
                                pop.CommunityPopulation = cvpop;
                                pop.HospitalPopulation = Decimal.Zero;
                            }
                            else if (sectAvail.AvailabilityHospital)
                            {
                                hvpop = (hospital == Decimal.Zero) ? total : hospital;
                                pop.Country = ctry;
                                pop.Year = (int)year;
                                pop.ATCClass = atcClass;
                                pop.Sector = sect;
                                pop.TotalPopulation = Decimal.Zero;
                                pop.CommunityPopulation = Decimal.Zero;
                                pop.HospitalPopulation = hvpop;
                            }
                            if (string.IsNullOrEmpty(pop.Country)) continue;
                            // Initialize or resize population array
                            if (noPop)
                            {
                                pops = new List<Population>(); // Initialize population list
                                noPop = false;
                            }
                            pops.Add(pop);
                        }
                    }
                    continue;
                }
                if (string.IsNullOrEmpty(atcClass) && sector!=null)
                {
                    // Similar logic as above for sector-specific population
                    foreach (var atcClassItem in AMUConstants.ATCClasses)
                    {
                        var pop = new Population();
                        atcClass = atcClassItem;

                        bool av = Utils.IsAmClassYearAvailability(atcClass, (int)year, avData);
                        if (!av) continue;

                        var atcClassAvail = Utils.GetATCClassYearAvailability(atcClass, (int)year, avData);
                        HealthSector hs = (HealthSector)sector;
                        var sectAvail = atcClassAvail[hs];

                        if (sectAvail.AvailabilityTotal)
                        {
                            pop.Country = ctry;
                            pop.Year = (int)year;
                            pop.ATCClass = atcClass;
                            pop.Sector = (HealthSector)sector;
                            pop.TotalPopulation = (Decimal)total;
                            pop.CommunityPopulation = Decimal.Zero;
                            pop.HospitalPopulation = Decimal.Zero;
                        }
                        else if (sectAvail.AvailabilityCommunity && sectAvail.AvailabilityHospital)
                        {
                            cvpop = (community == Decimal.Zero) ? total : community;
                            hvpop = (hospital == Decimal.Zero) ? total : hospital;

                            pop.Country = ctry;
                            pop.Year = (int)year;
                            pop.ATCClass = atcClass;
                            pop.Sector = (HealthSector)sector;
                            pop.TotalPopulation = Decimal.Zero;
                            pop.CommunityPopulation = cvpop;
                            pop.HospitalPopulation = hvpop;
                        }
                        else if (sectAvail.AvailabilityCommunity)
                        {
                            cvpop = (community == Decimal.Zero) ? total : community;
                            pop.Country = ctry;
                            pop.Year = (int)year;
                            pop.ATCClass = atcClass;
                            pop.Sector = (HealthSector)sector;
                            pop.TotalPopulation = Decimal.Zero;
                            pop.CommunityPopulation = cvpop;
                            pop.HospitalPopulation = Decimal.Zero;
                        }
                        else if (sectAvail.AvailabilityHospital)
                        {
                            hvpop = (hospital == Decimal.Zero) ? total : hospital;
                            pop.Country = ctry;
                            pop.Year = (int)year;
                            pop.ATCClass = atcClass;
                            pop.Sector = (HealthSector)sector;
                            pop.TotalPopulation = Decimal.Zero;
                            pop.CommunityPopulation = Decimal.Zero;
                            pop.HospitalPopulation = hvpop;
                        }

                        if (noPop)
                        {
                            pops = new List<Population>();
                            noPop = false;
                        }
                        pops.Add(pop);
                    }
                    continue;
                }
                if (!string.IsNullOrEmpty(atcClass) && sector==null)
                {
                    // Special case: ATC class is provided, but sector is not
                    if (!Utils.IsAmClassYearAvailability(atcClass, (int)year, avData))
                    {
                        // Equivalent to GoTo continueR
                        return false;
                    }
                    var amClassAvail = Utils.GetATCClassYearAvailability(atcClass, (int)year, avData);
                    var sectors = amClassAvail.Keys.ToList();
                    // This is the case where amClass is provided, but sector is missing
                    foreach (var sect in sectors)
                    {
                        var pop = new Population();
                        var sectAvail = amClassAvail[sect];

                        bool av = Utils.IsAmClassYearAvailability(atcClass, (int)year, avData);
                        if (!av) continue;



                        if (sectAvail.AvailabilityTotal)
                        {
                            pop.Country = ctry;
                            pop.Year = (int)year;
                            pop.ATCClass = atcClass;
                            pop.Sector = sect;
                            pop.TotalPopulation = total;
                            pop.CommunityPopulation = Decimal.Zero;
                            pop.HospitalPopulation = Decimal.Zero;
                        }
                        else if (sectAvail.AvailabilityCommunity && sectAvail.AvailabilityHospital)
                        {
                            cvpop = (community == Decimal.Zero) ? total : community;
                            hvpop = (hospital == Decimal.Zero) ? total : hospital;

                            pop.Country = ctry;
                            pop.Year = (int)year;
                            pop.ATCClass = atcClass;
                            pop.Sector = sect;
                            pop.TotalPopulation = Decimal.Zero;
                            pop.CommunityPopulation = cvpop;
                            pop.HospitalPopulation = hvpop;
                        }
                        else if (sectAvail.AvailabilityCommunity)
                        {
                            cvpop = (community == Decimal.Zero) ? total : community;
                            pop.Country = ctry;
                            pop.Year = (int)year;
                            pop.ATCClass = atcClass;
                            pop.Sector = sect;
                            pop.TotalPopulation = Decimal.Zero;
                            pop.CommunityPopulation = cvpop;
                            pop.HospitalPopulation = Decimal.Zero;
                        }
                        else if (sectAvail.AvailabilityHospital)
                        {
                            hvpop = (hospital == Decimal.Zero) ? total : hospital;
                            pop.Country = ctry;
                            pop.Year = (int)year;
                            pop.ATCClass = atcClass;
                            pop.Sector = sect;
                            pop.TotalPopulation = Decimal.Zero;
                            pop.CommunityPopulation = Decimal.Zero;
                            pop.HospitalPopulation = hvpop;
                        }
                        if (noPop)
                        {
                            pops = new List<Population>();
                            noPop = false;
                        }

                        pops.Add(pop);
                    }
                    continue;
                }
                else
                {
                    // Case: we are in a case where an ATCClass and a sector are given
                    
                    var pop = new Population
                    {
                        Country = ctry,
                        Year = (int)year,
                        ATCClass = atcClass,
                        Sector = (HealthSector)sector,
                        TotalPopulation = total,
                        CommunityPopulation = community,
                        HospitalPopulation = hospital
                    };

                    if (noPop)
                    {
                        pops = new List<Population>();
                        noPop = false;
                    }
                    pops.Add(pop);
                }
                continue;
            }
            foreach (var pop in pops)
            {
                // Check if the population entry already exists in popYears
                var existingPop = popYears.FirstOrDefault(p =>
                    p.Country == pop.Country &&
                    p.Year == pop.Year &&
                    p.ATCClass == pop.ATCClass &&
                    p.Sector == pop.Sector);

                if (existingPop != null)
                {
                    // Validate that we are not adding conflicting entries
                    if (existingPop.ATCClass == "ALL" || pop.ATCClass == "ALL" ||
                        existingPop.Sector == HealthSector.Total || pop.Sector == HealthSector.Total)
                    {
                        MessageBox.Show($"In {TemplateFormat.POPULATION_SHEETNAME} worksheet: Provide only one population figure per country, year, antimicrobial class, and sector.");
                        return false;
                    }
                }
                else
                {
                    popYears.Add(pop);
                }
            }
            return true;
        }


        public static bool ParseAvailability(Worksheet worksheet, int[] years,
                           out List<DataAvailability> availData)
        {
            availData = new List<DataAvailability>();  // Changed from Dictionary to List

            Range usedRange = worksheet.UsedRange;
            int rowWithData = Helper.Utils.GetRowsCountAvailabilityData(usedRange);
            // Bulk read entire sheet into a 2D object array for faster processing
            object[,] values = (object[,])usedRange.Value2;

            ErrorStatus errorStatus = new ErrorStatus();
            string variableErrors = "\n";

            for (int yIdx = 0; yIdx < years.Length; yIdx++)  // Iterate over years
            {
                int year = years[yIdx];

                // Iterate over rows in the worksheet
                for (int row = 2; row <= rowWithData; row++)  // Starting from row 2 (assuming first row is header)
                {
                    errorStatus.Reset();

                    // Parse required fields
                    string country = CommonParser.ParseCountryISO3(Convert.ToString(values[row, 1]), "Country", errorStatus, ref variableErrors, true);
                    if (errorStatus.Status > EntityStatus.OK)
                    {
                        if (string.IsNullOrEmpty(country))
                        {
                            MessageBox.Show($"In {TemplateFormat.AVAILABILITY_SHEETNAME} worksheet: Country {country} is not provided at row {row}.");
                        }
                        else {
                            MessageBox.Show($"In {TemplateFormat.AVAILABILITY_SHEETNAME} worksheet: Country {country} is not valid at row {row}.");
                        }
                        
                        return false;
                    }

                    string? atcClass = CommonParser.ParseATCClass(Convert.ToString(values[row, 2]), "ATC Class", errorStatus, ref variableErrors, true);
                    if (errorStatus.Status > EntityStatus.OK)
                    {
                        MessageBox.Show($"In {TemplateFormat.AVAILABILITY_SHEETNAME} worksheet: ATC Class {atcClass} is not valid at row {row}.");
                        return false;
                    }

                    HealthSector? sectorTmp = NAMUCommonParser.ParseHealthSector(Convert.ToString(values[row, 3]), "Sector", errorStatus, ref variableErrors, true);
                    if (sectorTmp == null)
                    {
                        MessageBox.Show($"In {TemplateFormat.AVAILABILITY_SHEETNAME} worksheet: Sector {sectorTmp} is not valid at row {row}.");
                        return false;
                    }
                    HealthSector sector = (HealthSector)sectorTmp;

                    // Constants 
                    const string UNK_AM_CLASS = "UNK_AM_CLASS";

                    // Check if the entry already exists in the list
                    var atcAvail = availData.FirstOrDefault(a => a.ATCClass == atcClass && a.Sector == sector && a.Year == year && a.Country == country);
                    var zatcAvail = availData.FirstOrDefault(a => a.ATCClass == UNK_AM_CLASS && a.Sector == sector && a.Year == year && a.Country == country);

                    // If the entry does not exist, create it and add it to the list
                    if (atcAvail == null)
                    {
                        atcAvail = new DataAvailability
                        {
                            ATCClass = atcClass,
                            Country = country,
                            Year = year,
                            Sector = sector
                        };
                        availData.Add(atcAvail);
                    }

                    if (zatcAvail == null)
                    {
                        zatcAvail = new DataAvailability
                        {
                            Country = country,
                            Year = year,
                            Sector = sector,
                            ATCClass = UNK_AM_CLASS
                        };
                        availData.Add(zatcAvail);
                    }

                    // Additional processing for levels (TOTAL, COMMUNITY, HOSPITAL)
                    for (int j = 1; j <= 3; j++)  // Loop for level processing (TOTAL, COMMUNITY, HOSPITAL)
                    {
                        errorStatus.Reset();
                        string levelCode = CommonParser.ParseString(Convert.ToString(values[row, 4]), "Level", errorStatus, ref variableErrors, true);
                        if (errorStatus.Status > EntityStatus.OK)
                        {
                            MessageBox.Show($"In {TemplateFormat.AVAILABILITY_SHEETNAME} worksheet: Level {levelCode} is not valid at row {row}.");
                            return false;
                        }
                        HealthLevel level;
                        try
                        {
                            level = HealthSectorLevelString.GetHealthLevelForString(levelCode);
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show($"In {TemplateFormat.AVAILABILITY_SHEETNAME} worksheet: Level {levelCode} is not valid at row {row}.");
                            return false;
                        }

                        // Assuming 'errorStatus' is your status object with the 'Status' property and 'ParseBoolean' returns a nullable boolean
                        errorStatus.Reset();  // Reset error status
                        int columnIndex = yIdx + TemplateFormat.AVAILABILITY_START_YEAR_COL_INDEX;  // Adjust the column index for the year
                        bool? tmpBool = CommonParser.ParseBoolean(Convert.ToString(values[row, columnIndex]), "Year", errorStatus, ref variableErrors, false);

                        // Check if the parsing returned a valid status
                        //if (errorStatus.Status > EntityStatus.OK)
                        //{
                        //    MessageBox.Show($"Value for year {year} is not valid in DataAvailability row {row}");
                        //    return false;  // Exit function if invalid data
                        //}

                        // Handle the case where tmpBool is null
                        bool avBool = tmpBool ?? false;  // If tmpBool is null, set avBool to false

                        // Assign the availability data to the corresponding level
                        switch (level)
                        {
                            case HealthLevel.Total:
                                atcAvail.AvailabilityTotal = avBool;
                                break;

                            case HealthLevel.Community:
                                atcAvail.AvailabilityCommunity = avBool;
                                break;

                            case HealthLevel.Hospital:
                                atcAvail.AvailabilityHospital = avBool;
                                break;
                        }
                    }
                }
            }

            return true;
        }


        public static int[] ParseAvailabilityYears(Range row)
        {
            // Get the total number of columns in the row
            int columnCount = row.Columns.Count;

            // Adjust starting index for years (based on your data structure)
            // Assuming years start from the 5th column

            // List to store years
            List<int> years = new List<int>();

            // Iterate through the year columns
            for (int i = TemplateFormat.AVAILABILITY_START_YEAR_COL_INDEX; i <= columnCount; i++)
            {
                // Read cell value (header)
                var cellValue = row.Worksheet.Cells[1, i].Value2;

                // Ensure the cell value is valid
                if (cellValue != null)
                {
                    // Parse value as integer
                    if (int.TryParse(cellValue.ToString().Trim(), out int year))
                    {
                        years.Add(year);
                    }
                }

            }

            return years.ToArray();
        }

    }


    public class ProductDataParser
    {

        #region Parse Consumption

        public static int[] ParsePackYears(Worksheet sheet)
        {
            List<int> years = new List<int>();

            // Determine the last column in the first row
            Range firstRow = sheet.Rows[1] as Range;
            int lColumn = firstRow.Cells[1, sheet.Columns.Count].End[XlDirection.xlToLeft].Column;

            int nbYears = lColumn - TemplateFormat.DATA_SHEET_PACKS_START_COL_IDX + 1;

            if (nbYears % 3 != 0)
            {
                MessageBox.Show(
                    "The number of years of data is not valid. Each year should be repeated three times: Total sector, Community sector, and Hospital sector even if you are not providing data for all three sectors.",
                    "Invalid Data",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                // Log error (you can replace this with your error handling logic)
                throw new InvalidOperationException("The number of years columns is not a multiple of 3.");
            }

            nbYears /= 3; // Each year is represented by three columns

            // Read years from the relevant columns
            for (int y = 0; y < nbYears; y++)
            {
                int year = Convert.ToInt32(
                    sheet.Cells[1, TemplateFormat.DATA_SHEET_PACKS_START_COL_IDX + (3 * y)].Value
                );
                years.Add(year);
            }

            return years.ToArray();
        }

        public static void ParsePackCons(Range drng, int[] years, int dLineNo,
                                   List<Product> products,
                                   List<ProductConsumption> productConsumptions,
                                   List<Population> populationData,
                                   List<DataAvailability> availabilityData,
                                   ErrorStatus es)
        {
            foreach (int year in years)
            {
                if (!availabilityData.Any(a => a.Year == year))
                {
                    continue;
                }
                var availData = availabilityData.Where(a => a.Year == year).ToList();
                products = SharedData.Products;

                if (products.Any(p => p.ProductLineNo == dLineNo))
                {
                    var pr = products.FirstOrDefault(p => p.ProductLineNo == dLineNo);
                    int lineNo = pr.ProductLineNo;
                    int seqNo = pr.SequenceNo;
                    // Check if ATC5 is null or empty and skip this product if so
                    if (string.IsNullOrEmpty(pr.ATC5.Code))
                    {
                        es.AddErrorMsgs($"Line {lineNo}: ATC5 code is null or empty. Skipping calculation of this lineNo.");
                        continue; // Skip to the next iteration of the year loop.
                    }
                    // Initialize ProductConsumption if it doesn't exist for the year and lineNo
                    var prodCons = productConsumptions.FirstOrDefault(pc => pc.Year == year && pc.LineNo == lineNo)
                                   ?? new ProductConsumption();

                    prodCons.ProductUniqueId = pr.UniqueId;
                    prodCons.LineNo = lineNo;
                    prodCons.Sequence = seqNo;
                    prodCons.ProductId = pr.ProductId;
                    prodCons.ATC5 = pr.ATC5.Code;
                    prodCons.AMClass = pr.AMClass;
                    prodCons.AtcClass = pr.ATCClass;
                    prodCons.Sector = pr.Sector;
                    prodCons.Year = year;
                    prodCons.AWaRe = pr.AWaRe;
                    prodCons.MEML = pr.MEML;

                    if (pr.GetValidate(Product.ROA_VALIDATION))
                    {
                        prodCons.Roa = pr.Roa.Code;
                        if (prodCons.Roa == "IS" || prodCons.Roa == "IP")
                        {
                            prodCons.Roa = "I";
                        }
                    }

                    prodCons.DPP = pr.NbDDD;

                    // Determine availability (refactored for list-based availability data)
                    string avAtcClass = null;
                    foreach (var atcClass in availData.Select(a => a.ATCClass).Distinct())
                    {
                        if (prodCons.ATC5.StartsWith(atcClass, StringComparison.OrdinalIgnoreCase))
                        {
                            avAtcClass = atcClass;
                            break;
                        }
                    }

                    if (!string.IsNullOrEmpty(avAtcClass))
                    {
                        var avData = availData.FirstOrDefault(a => a.ATCClass == avAtcClass && a.Sector == prodCons.Sector);

                        // Total Level
                        if (avData.AvailabilityTotal)
                        {
                            var dCell = drng.Cells[1, TemplateFormat.DATA_SHEET_PACKS_START_COL_IDX + (3 * Array.IndexOf(years, year))];
                            if (Decimal.TryParse(dCell?.Value?.ToString(), out Decimal packTotal))
                            {
                                prodCons.PKGConsumptionTotal = packTotal;
                            }
                            else if (string.IsNullOrEmpty(dCell?.Value?.ToString()))
                            {
                                prodCons.PKGConsumptionTotal = Decimal.Zero;
                            }
                            else
                            {
                                es.AddErrorMsgs($"Value for TOTAL_PACKAGE at line {dLineNo} is not valid");
                                return;
                            }
                            prodCons.AvailabilityTotal = true;
                        }

                        // Community Level
                        if (avData.AvailabilityCommunity)
                        {
                            var dCell = drng.Cells[1, TemplateFormat.DATA_SHEET_PACKS_START_COL_IDX + (3 * Array.IndexOf(years, year)) + 1];
                            if (Decimal.TryParse(dCell?.Value?.ToString(), out Decimal packComm))
                            {
                                prodCons.PKGConsumptionCommunity = packComm;
                            }
                            else if (string.IsNullOrEmpty(dCell?.Value?.ToString()))
                            {
                                prodCons.PKGConsumptionCommunity = Decimal.Zero;
                            }
                            else
                            {
                                es.AddErrorMsgs($"Value for COMMUNITY_PACKAGE at line {dLineNo} is not valid!");
                                return;
                            }

                            prodCons.AvailabilityCommunity = true;
                        }

                        // Hospital Level
                        if (avData.AvailabilityHospital)
                        {
                            var dCell = drng.Cells[1, TemplateFormat.DATA_SHEET_PACKS_START_COL_IDX + (3 * Array.IndexOf(years, year)) + 2];
                            if (Decimal.TryParse(dCell?.Value?.ToString(), out Decimal packHosp))
                            {
                                prodCons.PKGConsumptionHospital = packHosp;
                            }
                            else if (string.IsNullOrEmpty(dCell?.Value?.ToString()))
                            {
                                prodCons.PKGConsumptionHospital = Decimal.Zero;
                            }
                            else
                            {
                                es.AddErrorMsgs($"Value for HOSPITAL_PACKAGE at line {dLineNo} is not valid!");
                                return;
                            }
                            prodCons.AvailabilityHospital = true;
                        }

                        prodCons.CalculateDDD();

                        if (populationData.Any(p => p.Year == year))
                        {
                            // Find the population record for the given AMClass (fallback to "ALL" if not found)
                            var popAtcClass = populationData.Where(p => p.Year == year && p.ATCClass == prodCons.AtcClass).ToList();
                            if (popAtcClass.Count == 0)
                            {
                                popAtcClass = populationData.Where(p => p.Year == year && p.ATCClass == "ALL").ToList();
                            }

                            // If found, proceed with sector-specific population data (fallback to "GLO" if not found)
                            var popData = popAtcClass
                                          .Where(p => p.Sector == prodCons.Sector)
                                          .FirstOrDefault()
                                          ?? popAtcClass.Where(p => p.Sector == HealthSector.Total).FirstOrDefault();

                            if (popData != null)
                            {
                                prodCons.PopulationTotal = popData.TotalPopulation ?? Decimal.Zero;

                                if (prodCons.PopulationTotal != Decimal.Zero)
                                {
                                    prodCons.PopulationCommunity = prodCons.PopulationTotal;
                                    prodCons.PopulationHospital = prodCons.PopulationTotal;
                                }
                                else
                                {
                                    prodCons.PopulationCommunity = popData.CommunityPopulation ?? Decimal.Zero;
                                    prodCons.PopulationHospital = popData.HospitalPopulation ?? Decimal.Zero;
                                }

                                prodCons.CalculateDID();
                            }
                        }
                        // Ensure prodCons is added to the list
                        productConsumptions.Add(prodCons);
                    }
                }
            }
        }

        public static List<ProductConsumption> ParsePackages(
        int[] years,
        Worksheet dataSheet,
        List<Product> productData,
        List<Population> populationData,
        List<DataAvailability> availabilityData,
        ErrorStatus es)
        {

            // Call ParsePackData and return the result
            return ParsePackData(dataSheet, years, productData, populationData, availabilityData, es);
        }

        private static List<ProductConsumption> ParsePackData(Worksheet sheet, int[] years,
        List<Product> productData, List<Population> populationData,
        List<DataAvailability> availabilityData, ErrorStatus es)
        {
            int startRow = TemplateFormat.DATA_SHEET_START_DATA_ROW_IDX;
            int noOfProducts = Helper.Utils.GetProductRowCount();
            int maxRow = startRow + noOfProducts - 1; // Last row index

            Range yearsColRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 1]];
            Range lastCol = yearsColRange.End[Microsoft.Office.Interop.Excel.XlDirection.xlToRight];
            int lastColumn = lastCol.Column;

            // Bulk Read: Load entire data range at once
            Range dataRange = sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[maxRow, lastColumn]];

            object[,] values = (object[,])dataRange.Value2;

            var productConsumptionData = new List<ProductConsumption>();

            // 🚀 Process rows in memory
            for (int i = 0; i < noOfProducts; i++)
            {
                int excelRow = startRow + i; // Adjust for Excel row numbering
                ParsePackConsFromArray(values, years, i, excelRow, productData, productConsumptionData, populationData, availabilityData, es, lastColumn);
            }

            return productConsumptionData;
        }

        private static void ParsePackConsFromArray(object[,] values, int[] years, int index, int excelRow,
        List<Product> products, List<ProductConsumption> productConsumptionData,
        List<Population> populationData, List<DataAvailability> availabilityData, ErrorStatus es, int lastColumn)
        {
            foreach (int year in years)
            {
                if (!availabilityData.Any(a => a.Year == year))
                {
                    continue;
                }

                var availData = availabilityData
                        .Where(a => a.Year == year && (a.AvailabilityTotal || a.AvailabilityCommunity || a.AvailabilityHospital))
                        .ToList();

                // products = SharedData.Products;

                if (products.Any(p => p.ProductLineNo == excelRow))
                {
                    var pr = products.FirstOrDefault(p => p.ProductLineNo == excelRow);
                    int lineNo = pr.ProductLineNo;
                    int seqNo = pr.SequenceNo;

                    //// Check if ATC5 is null or empty and skip this product if so
                    //if (string.IsNullOrEmpty(pr.ATC5))
                    //{
                    //    es.AddErrorMsgs($"Line {lineNo}: ATC5 code is null or empty. Skipping calculation of this line.");
                    //    continue;
                    //}
                    // check product is valid and DPP is not None
                    if (!pr.IsProductValid() || pr.NbDDD==Decimal.Zero)
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
                    prodCons.ProductUniqueId = pr.UniqueId;
                    prodCons.Label = pr.Label;  
                    prodCons.Country = pr.Country;
                    prodCons.Year = year;
                    prodCons.Sector = pr.Sector;

                    prodCons.ATC5 = pr.ATC5.Code;
                    prodCons.AMClass = pr.AMClass;
                    prodCons.AtcClass = pr.ATCClass;
                    prodCons.Sector = pr.Sector;
                    prodCons.Year = year;
                    prodCons.AWaRe = pr.AWaRe;
                    prodCons.MEML = pr.MEML;
                    prodCons.Paediatric = pr.Paediatric;

                    prodCons.Roa = pr.Roa.Code;
                    if (prodCons.Roa == "IS" || prodCons.Roa == "IP")
                    {
                        prodCons.Roa = "I";
                    }

                    prodCons.DPP = pr.NbDDD;

                    // Determine availability (refactored for list-based availability data)
                    string avAtcClass = availData.Select(a => a.ATCClass)
                        .FirstOrDefault(atcClass => prodCons.ATC5.StartsWith(atcClass, StringComparison.OrdinalIgnoreCase));

                    if (!string.IsNullOrEmpty(avAtcClass) && availData.Any(a => a.ATCClass == avAtcClass && a.Sector == prodCons.Sector))
                    {
                        var avData = availData.FirstOrDefault(a => a.ATCClass == avAtcClass && a.Sector == prodCons.Sector);

                        if (avData != null)
                        {
                            // 🚀 Extract values from `values[,]` instead of accessing Excel cells
                            int baseColIndex = TemplateFormat.DATA_SHEET_PACKS_START_COL_IDX + (3 * Array.IndexOf(years, year));

                            prodCons.AvailabilityTotal = avData.AvailabilityTotal;
                            if (prodCons.AvailabilityTotal) { 
                                prodCons.PKGConsumptionTotal = ExtractConsumptionValue(values, index, baseColIndex, es, "TOTAL_PACKAGE", lineNo, lastColumn);
                            } else
                            {
                                prodCons.PKGConsumptionTotal = decimal.Zero;
                            }

                            prodCons.AvailabilityCommunity = avData.AvailabilityCommunity;
                            if (prodCons.AvailabilityCommunity)
                            {
                                prodCons.PKGConsumptionCommunity = ExtractConsumptionValue(values, index, baseColIndex + 1, es, "COMMUNITY_PACKAGE", lineNo, lastColumn);
                            }
                            else
                            {
                                prodCons.PKGConsumptionCommunity= Decimal.Zero;
                            }

                            prodCons.AvailabilityHospital = avData.AvailabilityHospital;
                            if (prodCons.AvailabilityHospital)
                            {
                                prodCons.PKGConsumptionHospital = ExtractConsumptionValue(values, index, baseColIndex + 2, es, "HOSPITAL_PACKAGE", lineNo, lastColumn);
                            } else
                            {
                                prodCons.PKGConsumptionHospital = Decimal.Zero;
                            }
                        }

                        prodCons.CalculateDDD();

                        if (populationData.Any(p => p.Year == year))
                        {
                            var popData = populationData.FirstOrDefault(p => p.Year == year && p.ATCClass == prodCons.AtcClass && p.Sector == prodCons.Sector) 
                                          ?? populationData.FirstOrDefault(p => p.Year == year && p.ATCClass == "ALL" && p.Sector == prodCons.Sector);

                            if (popData != null)
                            {
                                if (prodCons.AvailabilityTotal)
                                {
                                    prodCons.PopulationTotal = popData.TotalPopulation ?? decimal.Zero;
                                }
                                else
                                {
                                    if (prodCons.AvailabilityCommunity)
                                    {
                                        prodCons.PopulationCommunity = popData.CommunityPopulation ?? prodCons.PopulationTotal;
                                    }
                                    if (prodCons.AvailabilityHospital)
                                    {
                                        prodCons.PopulationHospital = popData.HospitalPopulation ?? prodCons.PopulationTotal;
                                    }
                                } 
                                    
                                prodCons.CalculateDID();
                            }
                        }

                        productConsumptionData.Add(prodCons);
                    }
                }
            }
        }

        private static Decimal ExtractConsumptionValue(object[,] values, int row, int col, ErrorStatus es, string label, int lineNo, int lastColumn)
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

        #endregion

        #region Parse Availability

        
        public static bool IsATCClassInAvailability(
        string atcClass,
        List<DataAvailability> availabilityData)
        {
            // Check if any record matches the AM class and has availability
            return availabilityData.Any(da =>
            da.ATCClass == atcClass &&
                (da.AvailabilityTotal || da.AvailabilityCommunity || da.AvailabilityHospital));
        }

        #endregion


        public static string ParseStringObject(string val, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (val == null)
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory";
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                    es.AddErrorMsgs(errMsg);
                }
                return string.Empty;
            }
            return val;
        }

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

        public static List<Product> ParseProducts()
        {
            List<Product> listProducts = new List<Product>();
            Worksheet dataSheet = Globals.ThisWorkbook.Sheets[TemplateFormat.DATA_SHEETNAME] as Worksheet;

            if (dataSheet == null) return listProducts; // Exit if sheet is missing

            int startRow = TemplateFormat.DATA_SHEET_START_DATA_ROW_IDX; // First data row
            int noOfProducts = Utils.GetProductRowCount(); // Get row count
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
                ProductLineNo = excelRow
            };

            try
            {
                ParseProductCountryISO3(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_COUNTRY_COL_IDX]), pr);
                ParseProductId(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_PRODUCT_ID_COL_IDX]), pr);
                ParseProductLabel(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_LABEL_COL_IDX]), pr);
                ParseProductPackSize(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_PACKSIZE_COL_IDX]), pr);
                ParseProductRoa(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_ROA_COL_IDX]), pr);
                ParseProductStrength(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_STRENGTH_COL_IDX]), pr);
                ParseProductStrengthUnit(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_STRENGTH_UNIT_COL_IDX]), pr);
                ParseProductConcentrationVolume(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_CONCENTRATION_COL_IDX]), pr);
                ParseProductVolume(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_VOLUME_COL_IDX]), pr);
                ParseProductAtc5(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_ATC5_COL_IDX]), pr);
                ParseProductSalt(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_SALT_COL_IDX]), pr);
                ParseProductCombination(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_COMBINATION_COL_IDX]), pr);
                ParseProductPaediatrics(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_PAEDIATRIC_COL_IDX]), pr);
                ParseProductForm(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_FORM_COL_IDX]), pr);
                ParseProductOrigin(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_ORIGIN_COL_IDX]), pr);
                ParseProductManufacturerCountryISO3(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_MAN_COUNTRY_COL_IDX]), pr);
                ParseProductMarketAuthHolder(values[index + 1, TemplateFormat.DATA_SHEET_MAH_COL_IDX], pr);
                ParseProductGenerics(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_GENERICS_COL_IDX]), pr);
                ParseProductName(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_PRODUCT_NAME_COL_IDX]), pr);
                ParseProductIngredients(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_INGREDIENTS_COL_IDX]), pr);
                ParseProductYearAuthorization(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_YEAR_AUTHORIZATION_COL_IDX]), pr);
                ParseProductYearWithdrawal(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_YEAR_WITHDRAWAL_COL_IDX]), pr);
                ParseProductSector(Convert.ToString(values[index + 1, TemplateFormat.DATA_SHEET_SECTOR_COL_IDX]), pr);
            }
            catch (Exception e)
            {
                MessageBox.Show($"An unexpected error occurred when parsing product at row {pr.ProductLineNo}: {e.Message}");
            }
            return pr;
        }


        private static string ParseCountryIso(string value)
        {
            if (value.Length != 3)
            {
                throw new ArgumentException($"Country ISO3 {value} value is invalid. It must be 3-letters code.");
            }
            return value;
        }

        private static string ParseNonEmptyString(string value)
        {
            if (String.IsNullOrEmpty(value))
            {
                throw new ArgumentException($"Empty value. It is mandatory.");
            }
            return value;
        }

        private static Decimal ParseDecimal(string value)
        {
            return Convert.ToDecimal(value);
        }

        private static AdministrationRoute ParseRoa(string value)
        {
            if(!ThisWorkbook.AdminRouteDataDict.ContainsKey(value))
            {
                throw new ArgumentException($"ROA value {value} is invalid.");
            }
            return ThisWorkbook.AdminRouteDataDict[value];
        }

        private static MeasureUnit ParseMeasureUnit(string value)
        {
            if (!ThisWorkbook.UnitDataDict.ContainsKey(value))
            {
                throw new ArgumentException($"MeasureUnit value {value} is invalid.");
            }
            return ThisWorkbook.UnitDataDict[value];
        }

        private static ATC ParseAtc5(string value)
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

        private static Salt ParseSalt(string value)
        {
            if (!ThisWorkbook.SaltDataDict.ContainsKey(value))
            {
                throw new ArgumentException($"Salt value {value} is invalid.");
            }
            return ThisWorkbook.SaltDataDict[value];
        }

        private static DDDCombination ParseCombination(string value)
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

        private static int ParseYear(string value)
        {
            int y = Convert.ToInt32(value);
            if (y<1970 || y> DateTime.Today.Year)
            {
                throw new ArgumentException($"Year {y} is invalid [1970-{DateTime.Today.Year}]");
            }
            return y;
        }


        public static void ParseProductCountryISO3(object cellValue, Product pr)
        {
            
            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.COUNTRY_FIELD,true,TemplateFormat.DATA_SHEET_COUNTRY_COL_IDX);

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
                    pr.SetField(Product.COUNTRY_FIELD, df);
                    pr.Country = co;
                }
                catch (ArgumentException e)
                {
                    errMsg = $"{Product.COUNTRY_FIELD}=>{e.Message}";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.COUNTRY_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.COUNTRY_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                pr.SetField(Product.COUNTRY_FIELD, df);
                
            }
        }

        public static void ParseProductId(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.PRODUCT_ID_FIELD, true, TemplateFormat.DATA_SHEET_PRODUCT_ID_COL_IDX);
            
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

        public static void ParseProductLabel(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.LABEL_FIELD, true, TemplateFormat.DATA_SHEET_LABEL_COL_IDX);

            string errMsg;

            // Check if the value is empty
            try
            {
                string value2 = ParseNonEmptyString(value);
                df.InputValue = value;
                df.Value = value2;
                df.IsValid = true;
                pr.SetField(Product.LABEL_FIELD, df);

            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.LABEL_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsMissing = true;
                df.IsValid = false;
                pr.SetField(Product.PRODUCT_ID_FIELD, df);
            }
        }

        public static void ParseProductPackSize(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Decimal> df = InitializeDataField<Decimal>(Product.PACKSIZE_FIELD, true, TemplateFormat.DATA_SHEET_PACKSIZE_COL_IDX);

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

        public static void ParseProductRoa(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<AdministrationRoute> df = InitializeDataField<AdministrationRoute>(Product.ROUTE_ADMIN_FIELD, true, TemplateFormat.DATA_SHEET_ROA_COL_IDX);

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
                    pr.SetField(Product.ROUTE_ADMIN_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.ROUTE_ADMIN_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.ROUTE_ADMIN_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.ROUTE_ADMIN_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                pr.SetField(Product.ROUTE_ADMIN_FIELD, df);
            }
        }

        public static void ParseProductStrength(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Decimal> df = InitializeDataField<Decimal>(Product.STRENGTH_FIELD, true, TemplateFormat.DATA_SHEET_STRENGTH_COL_IDX);

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
                    pr.SetField(Product.STRENGTH_FIELD, df);
                }
                catch (FormatException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.STRENGTH_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.STRENGTH_FIELD, df);
                }

            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.STRENGTH_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                pr.SetField(Product.STRENGTH_FIELD, df);
            }
        }

        public static void ParseProductStrengthUnit(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<MeasureUnit> df = InitializeDataField<MeasureUnit>(Product.STRENGTH_UNIT_FIELD, true, TemplateFormat.DATA_SHEET_STRENGTH_UNIT_COL_IDX);

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
                    pr.SetField(Product.STRENGTH_UNIT_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.STRENGTH_UNIT_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.STRENGTH_UNIT_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.STRENGTH_UNIT_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                pr.SetField(Product.STRENGTH_UNIT_FIELD, df);
            }
        }

        public static void ParseProductConcentrationVolume(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Decimal> df = InitializeDataField<Decimal>(Product.CONCENTRATION_VOLUME_FIELD, true, TemplateFormat.DATA_SHEET_CONCENTRATION_COL_IDX);

            string errMsg;

           
            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.CONCENTRATION_VOLUME_FIELD, df);
                return;
            }

            try
            {
                Decimal ps = ParseDecimal(value);
                df.InputValue = value;
                df.Value = ps;
                df.IsValid = true;
                pr.SetField(Product.CONCENTRATION_VOLUME_FIELD, df);
            }
            catch (FormatException)
            {
                // Add an error message to the product
                errMsg = $"{Product.CONCENTRATION_VOLUME_FIELD} value is invalid.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                pr.SetField(Product.CONCENTRATION_VOLUME_FIELD, df);
            }
        }

        public static void ParseProductVolume(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Decimal> df = InitializeDataField<Decimal>(Product.VOLUME_FIELD, true, TemplateFormat.DATA_SHEET_VOLUME_COL_IDX);

            string errMsg;


            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.VOLUME_FIELD, df);
                return;
            }

            try
            {
                Decimal ps = ParseDecimal(value);
                df.InputValue = value;
                df.Value = ps;
                df.IsValid = true;
                pr.SetField(Product.VOLUME_FIELD, df);
            }
            catch (FormatException)
            {
                // Add an error message to the product
                errMsg = $"{Product.VOLUME_FIELD} value is invalid.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                pr.SetField(Product.VOLUME_FIELD, df);
            }
        }

        public static void ParseProductAtc5(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<ATC> df = InitializeDataField<ATC>(Product.ATC5_FIELD, true, TemplateFormat.DATA_SHEET_ATC5_COL_IDX);

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
                    pr.SetField(Product.ATC5_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.ATC5_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.ATC5_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.ATC5_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                pr.SetField(Product.ATC5_FIELD, df);
            }
        }

        public static void ParseProductSalt(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<Salt> df = InitializeDataField<Salt>(Product.SALT_FIELD, true, TemplateFormat.DATA_SHEET_SALT_COL_IDX);

            string errMsg;

            if (string.IsNullOrEmpty(value)) { // set default salt XXXX
                var defaultSalt = "XXXX";
                Salt salt = ParseSalt(defaultSalt);
                df.InputValue = value;
                df.Value = salt;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.SALT_FIELD, df);
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
                pr.SetField(Product.SALT_FIELD, df);
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.SALT_FIELD} value is invalid.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                pr.SetField(Product.SALT_FIELD, df);
            }
        }

        public static void ParseProductCombination(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<DDDCombination> df = InitializeDataField<DDDCombination>(Product.COMBINATION_FIELD, true, TemplateFormat.DATA_SHEET_COMBINATION_COL_IDX);

            string errMsg;

            // Check if the value is empty

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.COMBINATION_FIELD, df);
                return;
            }
            
            try
            {
                string value2 = value.ToUpper();
                DDDCombination comb = ParseCombination(value2);
                df.InputValue = value;
                df.Value = comb;
                df.IsValid = true;
                pr.SetField(Product.COMBINATION_FIELD, df);
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.COMBINATION_FIELD} value is invalid.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                pr.SetField(Product.COMBINATION_FIELD, df);
            }
        }

        public static void ParseProductPaediatrics(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<YesNoUnknown> df = InitializeDataField<YesNoUnknown>(Product.PAEDIATRIC_FIELD, true, TemplateFormat.DATA_SHEET_PAEDIATRIC_COL_IDX);

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
                    pr.SetField(Product.PAEDIATRIC_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.PAEDIATRIC_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.PAEDIATRIC_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.PAEDIATRIC_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                pr.SetField(Product.PAEDIATRIC_FIELD, df);
            }
        }

        public static void ParseProductForm(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.FORM_FIELD, true, TemplateFormat.DATA_SHEET_FORM_COL_IDX);


            if(String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.FORM_FIELD, df);
            }
            else
            {
                df.InputValue = value;
                df.Value = value;
                df.IsValid = true;
                df.IsMissing = false;
                pr.SetField(Product.FORM_FIELD, df);
            }
        }

        public static void ParseProductOrigin(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.PRODUCT_ORIGIN_FIELD, true, TemplateFormat.DATA_SHEET_ORIGIN_COL_IDX);

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

        public static void ParseProductManufacturerCountryISO3(object cellValue, Product pr)
        {

            // Convert cellValue to string safely
            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.MANUFACTURER_COUNTRY_FIELD, false, TemplateFormat.DATA_SHEET_MAN_COUNTRY_COL_IDX);

            string errMsg;

            // Check if the value is empty

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.MANUFACTURER_COUNTRY_FIELD, df);
                return;
            }
            string value2 = value.ToUpper();
            
            try
            {
                string co = ParseCountryIso(value2);
                df.InputValue = value;
                df.Value = co;
                df.IsValid = true;
                pr.SetField(Product.MANUFACTURER_COUNTRY_FIELD, df);
            }
            catch (ArgumentException e)
            {
                errMsg = $"{Product.MANUFACTURER_COUNTRY_FIELD}=>{e.Message}";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                pr.SetField(Product.MANUFACTURER_COUNTRY_FIELD, df);
            }
        }

        public static void ParseProductGenerics(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<YesNoUnknown> df = InitializeDataField<YesNoUnknown>(Product.GENERICS_FIELD, true, TemplateFormat.DATA_SHEET_GENERICS_COL_IDX);

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

            DataField<string> df = InitializeDataField<string>(Product.PRODUCT_NAME_FIELD, true, TemplateFormat.DATA_SHEET_PRODUCT_NAME_COL_IDX);

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

        public static void ParseProductMarketAuthHolder(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.MARKET_AUTH_HOLDER_FIELD, true, TemplateFormat.DATA_SHEET_MAH_COL_IDX);

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.MARKET_AUTH_HOLDER_FIELD, df);
            }
            else
            {
                df.InputValue = value;
                df.Value = value;
                df.IsValid = true;
                df.IsMissing = false;
                pr.SetField(Product.MARKET_AUTH_HOLDER_FIELD, df);
            }
        }

        public static void ParseProductIngredients(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<string> df = InitializeDataField<string>(Product.INGREDIENTS_FIELD, true, TemplateFormat.DATA_SHEET_INGREDIENTS_COL_IDX);


            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsValid = true;
                df.IsMissing = true;
                pr.SetField(Product.INGREDIENTS_FIELD, df);
            }
            else
            {
                df.InputValue = value;
                df.Value = value;
                df.IsValid = true;
                df.IsMissing = false;
                pr.SetField(Product.INGREDIENTS_FIELD, df);
                if (value.Contains("+") || value.Contains("/"))
                {
                    var infMsg = $"{Product.INGREDIENTS_FIELD}=>if you have a fixed dose combination, please use comma to list each INN.";
                    pr.AddInfoMsg(infMsg);
                }
            }
        }

        public static void ParseProductYearAuthorization(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<int> df = InitializeDataField<int>(Product.YEAR_AUTHORIZATION_FIELD, true, TemplateFormat.DATA_SHEET_YEAR_AUTHORIZATION_COL_IDX);

            string errMsg;

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsMissing = true;
                df.IsValid = true;
                pr.SetField(Product.YEAR_AUTHORIZATION_FIELD, df);
            }
            else
            {
                try
                {
                    int y = ParseYear(value);
                    df.InputValue = value;
                    df.Value = y;
                    df.IsValid = true;
                    pr.SetField(Product.YEAR_AUTHORIZATION_FIELD, df);
                }
                catch (ArgumentException e)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.YEAR_AUTHORIZATION_FIELD}=>{e.Message}.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.YEAR_AUTHORIZATION_FIELD, df);
                }
                catch (FormatException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.YEAR_AUTHORIZATION_FIELD} is an invalid year.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.YEAR_AUTHORIZATION_FIELD, df);
                }
            }
        }

        public static void ParseProductYearWithdrawal(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<int> df = InitializeDataField<int>(Product.YEAR_WITHDRAWAL_FIELD, true, TemplateFormat.DATA_SHEET_YEAR_WITHDRAWAL_COL_IDX);

            string errMsg;

            if (String.IsNullOrEmpty(value))
            {
                df.InputValue = value;
                df.IsMissing = true;
                df.IsValid = true;
                pr.SetField(Product.YEAR_WITHDRAWAL_FIELD, df);
            }
            else
            {
                try
                {
                    int y = ParseYear(value);
                    df.InputValue = value;
                    df.Value = y;
                    df.IsValid = true;
                    pr.SetField(Product.YEAR_WITHDRAWAL_FIELD, df);
                }
                catch (ArgumentException e)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.YEAR_WITHDRAWAL_FIELD}=>{e.Message}.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.YEAR_WITHDRAWAL_FIELD, df);
                }
                catch (FormatException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.YEAR_WITHDRAWAL_FIELD} is an invalid year.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.YEAR_WITHDRAWAL_FIELD, df);
                }
            }
        }

        public static void ParseProductSector(object cellValue, Product pr)
        {

            string value = Convert.ToString(cellValue)?.Trim();

            DataField<HealthSector> df = InitializeDataField<HealthSector>(Product.SECTOR_FIELD, true, TemplateFormat.DATA_SHEET_SECTOR_COL_IDX);

            string errMsg;

            try
            {
                string value2 = ParseNonEmptyString(value);
                try
                {
                    string value3 = value2.ToUpper();
                    HealthSector hs = HealthSectorLevelString.GetHealthSectorForString(value3);
                    df.InputValue = value;
                    df.Value = hs;
                    df.IsValid = true;
                    pr.SetField(Product.SECTOR_FIELD, df);
                }
                catch (ArgumentException)
                {
                    // Add an error message to the product
                    errMsg = $"{Product.SECTOR_FIELD} value is invalid.";
                    pr.AddErrorMsgs(errMsg);
                    df.InputValue = value;
                    df.IsValid = false;
                    pr.SetField(Product.SECTOR_FIELD, df);
                }
            }
            catch (ArgumentException)
            {
                // Add an error message to the product
                errMsg = $"{Product.SECTOR_FIELD} is mandatory.";
                pr.AddErrorMsgs(errMsg);
                df.InputValue = value;
                df.IsValid = false;
                df.IsMissing = true;
                pr.SetField(Product.SECTOR_FIELD, df);
            }
        }
    }
}
