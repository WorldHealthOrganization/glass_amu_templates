// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Models;

namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>MeasureUnitListParser</c> provide parsing methods to load measure unit list definition.
    /// </summary>
    public class MeasureUnitListParser
    {
        /// <summary>
        /// Parse the measure unit list from a Excel Range.
        /// </summary>
        /// <param name="uListRange">The range that contains the measure unit list definition, including the header row.</param>
        /// <returns>A list of MeasureUnit objects</returns>
        public static List<MeasureUnit> ParseMeasureUnitList(Range uListRange)
        {
            List<MeasureUnit> uList = new List<MeasureUnit>();

            for (int row = 2; row <= uListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                string code = StringParser.ParseAndTrimAndUpper(uListRange.Cells[row, 1].Value);
                string unitFamilyStr = StringParser.ParseAndTrim(uListRange.Cells[row, 2].Value);
                MeasureUnitFamily unitFamily = MeasureUnitFamilyString.GetMeasureUnitFamilyFromString(unitFamilyStr);
                Decimal baseConv = Convert.ToDecimal(uListRange.Cells[row, 6].Value);

                uList.Add(new MeasureUnit(code, unitFamily, baseConv));
            }

            return uList;
        }
    }
}
