// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Models;
using HAMU_Template;
using HAMU_Template.Constants;
using HAMU_Template.Models;
using HAMU_Template.Models.Mappings;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace HAMU_Template.Helper
{
    public static class Utils
    {
        public static int GetRowsCountAvailabilityData(Excel.Range usedRange)
        {
            //Check for the cells data present in all the cells of column 1 -- 12
            int rowsWithData = 0;

            // Start from row 2 and check each row till the end..!
            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
                bool hasData = false;

                // Check all cells in columns 1 to 12..!
                for (int col = 1; col <= 12; col++)
                {
                    // If any cell is not empty, the row has data..!
                    if (usedRange.Cells[row, col].Value2 != null && usedRange.Cells[row, col].Value2.ToString() != "")
                    {
                        hasData = true;
                        break; // Stop checking further cells in this row..!
                    }
                }

                // If at least one cell had data, add 1 to the counter..!
                if (hasData)
                {
                    rowsWithData++;
                }
            }
            return rowsWithData;
        }

        public static int GetRowsCountHospitalStructure(Excel.Range usedRange)
        {
            //Check for the cells data present in all the cells of column 1 -- 3..!
            int rowsWithData = 0;

            // Start from row 2 and check each row till the end..!
            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
                bool hasData = false;

                // Check all cells in columns 3 to 8..!
                for (int col = 1; col <= 3; col++)
                {
                    // If any cell is not empty, the row has data..!
                    if (usedRange.Cells[row, col].Value2 != null && usedRange.Cells[row, col].Value2.ToString() != "")
                    {
                        hasData = true;
                        break; // Stop checking further cells in this row..!
                    }
                }

                // If at least one cell had data, add 1 to the counter..!
                if (hasData)
                {
                    rowsWithData++;
                }
            }
            return rowsWithData;

        }
        public static int GetRowsCountHospitalActivityData(Excel.Range usedRange)
        {

            //Check for the cells data present in all the cells of column 1 -- 6..!
            int rowsWithData = 0;

            // Start from row 2 and check each row till the end..!
            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
                bool hasData = false;

                // Check all cells in columns 1 to 6..!
                for (int col = 1; col <= 6; col++)
                {
                    // If any cell is not empty, the row has data..!
                    if (usedRange.Cells[row, col].Value2 != null && usedRange.Cells[row, col].Value2.ToString() != "")
                    {
                        hasData = true;
                        break; // Stop checking further cells in this row..!
                    }
                }

                // If at least one cell had data, add 1 to the counter..!
                if (hasData)
                {
                    rowsWithData++;
                }
            }
            return rowsWithData;
        }

        public static int GetRowsCountProductData(Excel.Range usedRange)
        {
            // Check for the cells data present in all the cells of column 3 -- 10..!
            int rowsWithData = 0;

            // Start from row 3 and check each row till the end..!
            for (int row = 3; row <= usedRange.Rows.Count; row++)
            {
                bool hasData = false; //false initially..!

                // Check all cells in columns 3 to 10..!
                for (int col = 3; col <= 10; col++)
                {
                    // If any cell is not empty, the row has data..!
                    if (usedRange.Cells[row, col].Value2 != null && usedRange.Cells[row, col].Value2.ToString() != "")
                    {
                        hasData = true;
                        break; // Stop checking further cells in this row..!
                    }
                }

                // If at least one cell had data, add 1 to the counter..!
                if (hasData)
                {
                    rowsWithData++;
                }
            }
            return rowsWithData;
        }

        public static int GetRowsCountSubstanceData(Excel.Range usedRange)
        {
            //Check for the cells data present in all the cells of column 3 -- 8..!
            int rowsWithData = 0;

            // Start from row 3 and check each row till the end..!
            for (int row = 3; row <= usedRange.Rows.Count; row++)
            {
                bool hasData = false;

                // Check all cells in columns 3 to 8..!
                for (int col = 3; col <= 8; col++)
                {
                    // If any cell is not empty, the row has data..!
                    if (usedRange.Cells[row, col].Value2 != null && usedRange.Cells[row, col].Value2.ToString() != "")
                    {
                        hasData = true;
                        break; // Stop checking further cells in this row..!
                    }
                }

                // If at least one cell had data, add 1 to the counter..!
                if (hasData)
                {
                    rowsWithData++;
                }
            }
            return rowsWithData;
        }

        public static int GetActualColumnForYears(Excel.Range ws, int startIndex, Excel.Worksheet header)
        {
            int actualCount = 0;

            if (ws == null)
            {
                return actualCount;
            }

            //Parse them as a string and then parse in double..!
            for (int colIndex = startIndex; colIndex <= ws.Columns.Count; colIndex++)
            {
                var yearCellValue = header.Cells[1, colIndex].Value2;
                if (yearCellValue == null)
                {
                    break;
                }
                actualCount++;
            }

            return actualCount;
        }

        public static bool IsATCClassInAvailability(string atcClass, List<Availability> availabilityData)
        {
            Availability res = availabilityData
                .FirstOrDefault<Availability>(m => m.ATCClass == atcClass);
            return res != null;
        }
    }
}
