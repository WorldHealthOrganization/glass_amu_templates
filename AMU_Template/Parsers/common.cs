// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;

namespace AMU_Template.Parsers
{
    public class StringParser
    {
        public static string ParseAndTrim(dynamic value)
        { 
            if (String.IsNullOrEmpty(value))
            {
                return value;
            }
            return Convert.ToString(value).Trim();
        }

        public static string ParseAndTrimAndUpper(dynamic value)
        {
            if (String.IsNullOrEmpty(value))
            {
                return value;
            }
            return StringParser.ParseAndTrim(value).ToUpper();
        }
    }
}
