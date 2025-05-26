// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Helpers;
using AMU_Template.Models;

namespace AMU_Template.Parsers
{
    /// <summary>
    /// Class <c>ConversionFactorListParser</c> provide parsing methods to load conversion factor list definition.
    /// </summary>
    public class ConversionFactorListParser
    {
        /// <summary>
        /// Parse the conversion factor list from a Excel Range.
        /// </summary>
        /// <param name="cfListRange">The range that contains the conversion factor list definition, including the header row.</param>
        /// <returns>A list of ConversionFactor objects</returns>
        public static List<ConversionFactor> ParseConversionFactorList(Range cfListRange, Dictionary<string, ATC> atc5_dict, Dictionary<string, AdministrationRoute> roa_dict, Dictionary<string, Salt> salt_dict, Dictionary<string, MeasureUnit> unit_dict)
        {
            List<ConversionFactor> cfList = new List<ConversionFactor>();

            for (int row = 2; row <= cfListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                ConversionFactor cf = new ConversionFactor();
                string atcCode = cfListRange.Cells[row, 2].Value;
                cf.ATC5 = atc5_dict[atcCode.Trim().ToUpper()];
                string roaCode = cfListRange.Cells[row, 3].Value;
                cf.ROA = roa_dict[roaCode.Trim().ToUpper()];
                string saltCode = cfListRange.Cells[row, 4].Value;
                if (string.IsNullOrEmpty(saltCode))
                {
                    saltCode = "XXXX";
                }
                cf.Salt = salt_dict[saltCode.Trim().ToUpper()];
                string fu = cfListRange.Cells[row, 5].Value;
                cf.UnitFrom = unit_dict[fu.Trim().ToUpper()];
                string tu = cfListRange.Cells[row, 6].Value;
                cf.UnitTo = unit_dict[tu.Trim().ToUpper()];
                cf.Factor = Convert.ToDecimal(cfListRange.Cells[row, 7].Value);
                cf.ARS = ARSHelper.GenerateARSFromATC5ROASalt(cf.ATC5.Code, cf.ROA.Code, cf.Salt.Code);
                cfList.Add(cf);
            }

            return cfList;
        }
    }
}
