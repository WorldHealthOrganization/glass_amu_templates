// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Constants;
using AMU_Template.Models;

namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>AwareListParser</c> provide parsing methods to load ATC list definition.
    /// </summary>
    public class MEMLListParser
    {

        public static string[] defaultRoas = { "O", "P", "R", "I", "IS", "IP" };

        /// <summary>
        /// Parse the mEML list from a Excel Range.
        /// </summary>
        /// <param name="emlListRange">The range that contains the mEML list definition, including the header row.</param>
        /// <returns>A list of MEML objects</returns>
        public static List<MEML> ParseMEMLList(Range emlListRange)
        {
            List<MEML> emlList = new List<MEML>();

            for (int row = 2; row <= emlListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                var atc5 = StringParser.ParseAndTrimAndUpper(emlListRange.Cells[row, 1].Value);
                string emlTxt = StringParser.ParseAndTrimAndUpper(emlListRange.Cells[row, 3].Value);
                var inEML = YesNoNAString.GetYesNoNAFromString(emlTxt);
                var roa = StringParser.ParseAndTrimAndUpper(emlListRange.Cells[row, 2].Value);
                if (String.IsNullOrEmpty(roa))
                {
                    for (int i = 0; i < defaultRoas.Length; i++)
                    {
                        emlList.Add(new MEML(atc5, defaultRoas[i], inEML));
                    }
                }
                else
                {
                    emlList.Add(new MEML(atc5, roa, inEML));
                }
            }

            return emlList;
        }
    }
}
