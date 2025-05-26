// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Parsers;
using NAMU_Template.Models;

namespace NAMU_Template.Data_Parsing
{
    public class CountryListParser
    {
        /// <summary>
        /// Parse the Country list from a Excel Range.
        /// </summary>
        /// <param name="cntrListRange">The range that contains the Country list definition, including the header row.</param>
        /// <returns>A list of Country objects</returns>
        public static List<Country> ParseCountryList(Range cntryListRange)
        {

            List<Country> cntryList = new List<Country>();

            for (int row = 2; row <= cntryListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                Country cntry = new Country
                {
                    Code = StringParser.ParseAndTrimAndUpper(cntryListRange.Cells[row, 1].Value),
                    ShortName = StringParser.ParseAndTrim(cntryListRange.Cells[row, 2].Value),
                    FormalName = StringParser.ParseAndTrim(cntryListRange.Cells[row, 3].Value),
                };
                cntryList.Add(cntry);
            }

            return cntryList;
        }
    }
}
