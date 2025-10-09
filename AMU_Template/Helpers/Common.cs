// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Models;
using System;
using System.Collections.Generic;

namespace AMU_Template.Helpers
{
    public static class ARSHelper
    {
        public static string GenerateARSFromATC5ROASalt(string ATC5, string ROA, string Salt)
        {
            return $"{ATC5}_{ROA}_{Salt}";
        }
    }

    public static class ATCHelper
    {
        public static ATC GetATCParent(ATC child, Dictionary<string, ATC> ATCDataDict)
        {
            var parent_code = "";
            switch (child.Level)
            {
                case 5:
                    parent_code = child.Code.Substring(0, 5);
                    break;
                case 4:
                    parent_code = child.Code.Substring(0, 4);
                    break;
                case 3:
                    parent_code = child.Code.Substring(0, 3);
                    break;
                case 2:
                    parent_code = child.Code.Substring(0, 1);
                    break;
                default:
                    throw new Exception($"ATC {child.Code} is already at the ATC level 1.");
            }
            if (!ATCDataDict.ContainsKey(parent_code))
            {
                throw new Exception($"ATC {parent_code} does not exist.");
            }
            ATC parent = ATCDataDict[parent_code];
            return parent;
        }
    }

    public static class DataHelper
    {
        public static object[,] ConvertTo2DArray(object[][] jaggedArray)
        {
            int rows = jaggedArray.Length;
            int cols = jaggedArray[0].Length;
            object[,] result = new object[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    result[i, j] = jaggedArray[i][j];
                }
            }

            return result;
        }
    }
}
     
