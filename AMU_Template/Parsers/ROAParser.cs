// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Models;




namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>RoaListParser</c> provide parsing methods to load the administration route list definition.
    /// </summary>
    public class RoaListParser
    {
        /// <summary>
        /// Parse the ATC list from a Excel Range.
        /// </summary>
        /// <param name="roaListRange">The range that contains the administration route list definition, including the header row.</param>
        /// <returns>A list of ATC objects</returns>
        public static List<AdministrationRoute> ParseRoaList(Range roaListRange)
        {

            List<AdministrationRoute> roaList = new List<AdministrationRoute>();

            for (int row = 2; row <= roaListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                string code = StringParser.ParseAndTrimAndUpper(roaListRange.Cells[row, 1].Value);
                string name = StringParser.ParseAndTrim(roaListRange.Cells[row, 2].Value);
                roaList.Add(new AdministrationRoute(code, name));
            }

            // Add Undefined Route for Z99ZZ99
            roaList.Add(new AdministrationRoute("X", "UNDEFINED"));

            return roaList;
        }
    }
}
