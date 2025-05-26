// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace AMU_Template.Helpers
{
    public static class ARSHelper
    {
        public static string GenerateARSFromATC5ROASalt(string ATC5, string ROA, string Salt)
        {
            return $"{ATC5}_{ROA}_{Salt}";
        }
    }
}
