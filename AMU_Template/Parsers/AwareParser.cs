// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>AwareListParser</c> provide parsing methods to load ATC list definition.
    /// </summary>
    public class AwareListParser
    {
        public static string[] defaultRoas = { "O", "P", "R", "I", "IS", "IP" };

        /// <summary>
        /// Parse the Aware list from a Excel Range.
        /// </summary>
        /// <param name="awrListRange">The range that contains the Aware list definition, including the header row.</param>
        /// <returns>A list of Aware objects</returns>
        /// 
        public static List<Aware> ParseAwareList(Range awrListRange)
        {
            List<Aware> awrList = new List<Aware>();

            for (int row = 2; row <= awrListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                
                var atc5 = StringParser.ParseAndTrimAndUpper(awrListRange.Cells[row, 1].Value);
                var awr = StringParser.ParseAndTrimAndUpper(awrListRange.Cells[row, 3].Value);
                var roa = StringParser.ParseAndTrimAndUpper(awrListRange.Cells[row, 2].Value);
                if (String.IsNullOrEmpty(roa))
                {
                    for (int i = 0; i < defaultRoas.Length; i++)
                    {
                        awrList.Add(new Aware(atc5, defaultRoas[i], awr));
                    }           
                }
                else
                {
                    awrList.Add(new Aware(atc5, roa, awr));
                }
            }

            return awrList;
        }
    }
}
