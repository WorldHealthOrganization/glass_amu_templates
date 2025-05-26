// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Models;

namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>DDDCombinationListParser</c> provide parsing methods to load DDD combination list definition.
    /// </summary>
    public class DDDCombinationListParser
    {
        /// <summary>
        /// Parse the conversion factor list from a Excel Range.
        /// </summary>
        /// <param name="cListRange">The range that contains the DDD combination list definition, including the header row.</param>
        /// <returns>A list of ConversionFactor objects</returns>
        public static List<DDDCombination> ParseDDDCombinationList(Range cListRange, Dictionary<string, ATC> atc5_dict, Dictionary<string, AdministrationRoute> roa_dict, Dictionary<string, MeasureUnit> unit_dict)
        {
            List<DDDCombination> cList = new List<DDDCombination>();

            for (int row = 2; row <= cListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                DDDCombination comb = new DDDCombination();
                comb.Code = StringParser.ParseAndTrimAndUpper(cListRange.Cells[row, 1].Value);
                string atcCode = StringParser.ParseAndTrimAndUpper(cListRange.Cells[row, 3].Value);
                comb.ATC5 = atc5_dict[atcCode];
                comb.Form = StringParser.ParseAndTrim(cListRange.Cells[row, 4].Value);
                string roaCode = StringParser.ParseAndTrimAndUpper(cListRange.Cells[row, 5].Value);
                comb.ROA = roa_dict[roaCode];
                comb.UnitDose = StringParser.ParseAndTrim(cListRange.Cells[row, 6].Value);
                comb.DDDValue = Convert.ToDecimal(cListRange.Cells[row, 7].Value);
                string u = StringParser.ParseAndTrimAndUpper(cListRange.Cells[row, 8].Value);
                comb.DDDUnit = unit_dict[u];
                comb.Info = StringParser.ParseAndTrim(cListRange.Cells[row, 9].Value);
                comb.Examples = StringParser.ParseAndTrim(cListRange.Cells[row, 10].Value);
                cList.Add(comb);
            }

            return cList;
        }
    }
}
