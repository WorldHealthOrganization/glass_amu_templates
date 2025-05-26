// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using AMU_Template.Parsers;
using NAMU_Template.Models;

namespace NAMU_Template.Data_Parsing
{
    public class ProductOriginListParser
    {
        /// <summary>
        /// Parse the Product Origin list from a Excel Range.
        /// </summary>
        /// <param name="origListRange">The range that contains the ProductOrigin list definition, including the header row.</param>
        /// <returns>A list of ProductOrigin objects</returns>
        public static List<ProductOrigin> ParseProductOriginList(Range origListRange)
        {

            List<ProductOrigin> origList = new List<ProductOrigin>();

            for (int row = 2; row <= origListRange.Rows.Count; row++) // Start from row 2 to skip the header..!
            {
                ProductOrigin pOrig = new ProductOrigin
                {
                    Code = StringParser.ParseAndTrimAndUpper(origListRange.Cells[row, 1].Value),
                    Name = StringParser.ParseAndTrim(origListRange.Cells[row, 2].Value)
                };
                origList.Add(pOrig);
            }

            return origList;
        }
    }
}
