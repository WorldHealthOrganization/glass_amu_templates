// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using NAMU_Template.Constants;
using NAMU_Template.Models;
using AMU_Template.Models;

namespace NAMU_Template.Helper
{
    public static class Utils
    {
        public static int GetRowsCountAvailabilityData(Excel.Range usedRange)
        {
            // Check for the cell
            int rowWithData = 0;


            // Start from row 2 and check each row till the end..
            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
                bool hasData = false;

                // check for the cells from c1 to c14
                for (int col = 1; col <= 14; col++)
                {
                    // if the cell is not empty the cell has data
                    if (usedRange.Cells[row, col].Value2 != null && usedRange.Cells[row, col].Value2.ToString() != "")
                    {
                        hasData = true;
                        break;
                    }
                }

                // if at least one cell has data, add 1 to the counter
                if (hasData)
                {
                    rowWithData++;
                }
            }
            return rowWithData;
        }

        public static int GetRowsCountPopulationData(Excel.Range usedRange)
        {
            int rowWithData = 0;


            // start from row 2and check each row till the end
            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
                bool hasData = false;

                // check for the cells from c1 to c4
                for (int col = 1; col <= 7; col++)
                {
                    // if the cell is not empty the cell has data
                    if (usedRange.Cells[row, col].Value2 != null && usedRange.Cells[row, col].Value2.ToString() != "")
                    {
                        hasData = true;
                        break;
                    }
                }

                // if at least one cell has data, add 1 to the counter
                if (hasData)
                {
                    rowWithData++;
                }
            }
            return rowWithData;
        }
        //Utils
        public static int GetProductRowCount()
        {
            int lineNo = Constants.START_DATA_ROW;
            int count = 0;
            bool doIter = true;

            //Excel.Worksheet dataSheet = Globals.ThisWorkbook.Sheets[SheetName.ProductData];
            Excel.Worksheet dataSheet = Globals.ThisWorkbook.Sheets[Constants.DATA_SHEET];

            while (doIter)
            {
                var cellValue = dataSheet.Range["C" + lineNo].Value;

                if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                {
                    doIter = false;
                }
                else
                {
                    count++;
                    lineNo++;
                }
            }

            return count;
        }
        public static bool IsEmptyStringCellObject(object cellValue)
        {
            return cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString());
        }

        public static bool IsEmptyStringCell(Excel.Range cell)
        {
            var cellValue = cell.Value;

            // Check if cell value is empty, null, or consists of only whitespace
            return cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString());
        }

        public static bool IsAmClassYearAvailability(string atcClass, int year, List<DataAvailability> availabilityData)
        {
            // Filter data for the given year and amClass
            var filteredData = availabilityData
                .Where(data => data.Year == year && data.ATCClass == atcClass)
                .ToList();

            // If no data is found for the given year and amClass, return false
            if (!filteredData.Any())
            {
                return false;
            }

            // Iterate through the filtered data and check availability
            foreach (var data in filteredData)
            {
                if (data.AvailabilityTotal || data.AvailabilityCommunity || data.AvailabilityHospital)
                {
                    return true; // If any availability is true, return true
                }
            }

            return false; // No availability found
        }

        public static Dictionary<HealthSector, DataAvailability> GetATCClassYearAvailability(
        string atcClass,
        int year,
        List<DataAvailability> availabilityData)
        {
            // Filter the availability data for the given AM class and year
            var filteredData = availabilityData
                .Where(da => da.ATCClass == atcClass && da.Year == year)
                .ToList();

            // Return a dictionary grouped by sector
            return filteredData
                .ToDictionary(da => da.Sector, da => da);
        }

        //public static ATC GetATCParent(ATC child)
        //{    
        //    var parent_code = "";
        //    switch(child.Level)
        //    {
        //        case 5:
        //            parent_code = child.Code.Substring(0,5);
        //            break;
        //        case 4:
        //            parent_code = child.Code.Substring(0, 4);
        //            break;
        //        case 3:
        //            parent_code = child.Code.Substring(0, 3);
        //            break;
        //        case 2:
        //            parent_code = child.Code.Substring(0, 1);
        //            break;
        //        default:
        //            throw new Exception($"ATC {child.Code} is already at the ATC level 1.");
        //    }
        //    ATC parent = ThisWorkbook.ATCDataDict[parent_code];
        //    return parent;
        //}
    }
}
