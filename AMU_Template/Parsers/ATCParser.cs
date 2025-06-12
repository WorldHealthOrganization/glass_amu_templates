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
                if (atc.Code == "Z99ZZ99")
                {
                    processZ99ATCs(atc, atcList);
                }
            }

            return atcList;
        }

        private static void processZ99ATCs(ATC atc5, List<ATC> atcList)
        {
            // process ATC4
            ATC atc4 = new ATC();
            atc4.Code = atc5.Code.Substring(0, 5);
            atc4.Name = atc5.Name;
            atc4.Level = 4;
            atc4.ATCClass = atc5.ATCClass;
            atc4.AMClass = atc5.AMClass;
            atcList.Add(atc4);

            // process ATC3
            ATC atc3 = new ATC();
            atc3.Code = atc5.Code.Substring(0, 4);
            atc3.Name = atc5.Name;
            atc3.Level = 4;
            atc3.ATCClass = atc5.ATCClass;
            atc3.AMClass = atc5.AMClass;
            atcList.Add(atc3);

            // process ATC2
            ATC atc2 = new ATC();
            atc2.Code = atc5.Code.Substring(0, 3);
            atc2.Name = atc5.Name;
            atc2.Level = 4;
            atc2.ATCClass = atc5.ATCClass;
            atc2.AMClass = atc5.AMClass;
            atcList.Add(atc2);

            // process ATC2
            ATC atc1 = new ATC();
            atc1.Code = atc5.Code.Substring(0, 1);
            atc1.Name = atc5.Name;
            atc1.Level = 4;
            atc1.ATCClass = atc5.ATCClass;
            atc1.AMClass = atc5.AMClass;
            atcList.Add(atc1);
        }
    }
}
