// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Models;

namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>ATCListParser</c> provide parsing methods to load the ATC list definition.
    /// </summary>
    public class ATCListParser
    {
        /// <summary>
        /// Parse the ATC list from a Excel Range.
        /// </summary>
        /// <param name="atcListRange">The range that contains the ATC list definition, including the header row.</param>
        /// <returns>A list of ATC objects</returns>
        public static List<ATC> ParseATCList(Range atcListRange)
        {

            List<ATC> atcList = new List<ATC>();

            for (int row = 2; row <= atcListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                ATC atc = new ATC(); 
                atc.Code = StringParser.ParseAndTrimAndUpper(atcListRange.Cells[row, 1].Value);
                atc.Name = StringParser.ParseAndTrim(atcListRange.Cells[row, 2].Value);
                atc.Level = Convert.ToInt32(atcListRange.Cells[row, 3].Value);
                atc.ATCClass= StringParser.ParseAndTrimAndUpper(atcListRange.Cells[row, 4].Value);
                atc.AMClass = StringParser.ParseAndTrimAndUpper(atcListRange.Cells[row, 5].Value);
                atcList.Add(atc);
            }

            return atcList;
        }
    }
}
