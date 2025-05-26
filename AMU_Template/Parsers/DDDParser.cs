// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Models;

namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>DDDListParser</c> provide parsing methods to load DDD list definition.
    /// </summary>
    public class DDDListParser
    {
        /// <summary>
        /// Parse the conversion factor list from a Excel Range.
        /// </summary>
        /// <param name="dListRange">The range that contains the DDD list definition, including the header row.</param>
        /// <returns>A list of ConversionFactor objects</returns>
        public static List<DDD> ParseDDDList(Range dListRange, Dictionary<string, ATC> atc5_dict, Dictionary<string, AdministrationRoute> roa_dict, Dictionary<string, Salt> salt_dict, Dictionary<string, MeasureUnit> unit_dict)
        {
            List<DDD> dList = new List<DDD>();

            for (int row = 2; row <= dListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                
                
                string atcCode = dListRange.Cells[row, 2].Value;
                ATC atc5 = atc5_dict[atcCode.Trim().ToUpper()];
                string roaCode = StringParser.ParseAndTrimAndUpper(dListRange.Cells[row, 3].Value);
                AdministrationRoute roa = roa_dict[roaCode];
                string saltCode = StringParser.ParseAndTrimAndUpper(dListRange.Cells[row, 4].Value);
                if (string.IsNullOrEmpty(saltCode))
                {
                    saltCode = "XXXX";
                }
                Salt salt = salt_dict[saltCode];
                Decimal dddValue = Convert.ToDecimal(dListRange.Cells[row, 5].Value);
                string u = StringParser.ParseAndTrimAndUpper(dListRange.Cells[row, 6].Value);
                MeasureUnit dddUnit = unit_dict[u];
                Decimal dddStdValue = Convert.ToDecimal(dListRange.Cells[row, 7].Value);
                string notes = StringParser.ParseAndTrim(dListRange.Cells[row, 8].Value);


                DDD ddd = new DDD(atc5, roa, salt, dddValue, dddUnit, dddStdValue, notes);
                dList.Add(ddd);
            }

            return dList;
        }
    }
}
