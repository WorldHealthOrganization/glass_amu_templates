// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Models;

namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>SaltListParser</c> provide parsing methods to load the salt list definition.
    /// </summary>
    public class SaltListParser
    {
        /// <summary>
        /// Parse the Salt list from a Excel Range.
        /// </summary>
        /// <param name="saltListRange">The range that contains the salt list definition, including the header row.</param>
        /// <returns>A list of ATC objects</returns>
        public static List<Salt> ParseSaltList(Range saltListRange)
        {

            List<Salt> saltList = new List<Salt>();
            List<string> atc5List = new List<string>();

            for (int row = 2; row <= saltListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                string code = StringParser.ParseAndTrimAndUpper(saltListRange.Cells[row, 1].Value);
                string name = StringParser.ParseAndTrim(saltListRange.Cells[row, 2].Value);
                string atc5sRaw = StringParser.ParseAndTrimAndUpper(saltListRange.Cells[row, 3].Value);
                if (!String.IsNullOrEmpty(atc5sRaw))
                {
                    var atc5s = atc5sRaw.Split(',');
                    atc5List = atc5s.ToList();

                }
                string? info = StringParser.ParseAndTrim(saltListRange.Cells[row, 4].Value);
                saltList.Add(new Salt(code, name, info, atc5List));
            }

            return saltList;
        }
    }
}
