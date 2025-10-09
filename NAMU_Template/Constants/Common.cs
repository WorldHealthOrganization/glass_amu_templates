// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace NAMU_Template.Constants
{
    public enum HealthSector
    {
        Public,
        Private,
        Total
    }

    public enum HealthLevel
    {
        Hospital,
        Community,
        Total
    }

    public class HealthSectorLevelString
    {
        public static HealthSector GetHealthSectorForString(string hsStr)
        {
            string s2 = hsStr.ToUpper().Trim();
            if (s2 == "PUBLIC" || s2 == "PUB")
            {
                return HealthSector.Public;
            }
            if (s2 == "PRIVATE" || s2 == "PRI")
            {
                return HealthSector.Private;
            }
            if (s2 == "TOTAL" || s2 == "TOT" || s2 == "GLOBAL" || s2 == "GLO")
            {
                return HealthSector.Total;
            }
            throw new ArgumentException("Health sector is not valid", hsStr);
        }

        public static string GetStringForHealthSector(HealthSector hs)
        {
            string value = "";
            switch(hs)
            {
                case HealthSector.Public:
                    value = "PUB";
                    break;
                case HealthSector.Private:
                    value = "PRI";
                    break;
                case HealthSector.Total:
                    value = "GLO";
                    break;
                default:
                    break;
            }
            return value;
        }


        public static HealthLevel GetHealthLevelForString(string hlStr)
        {
            string s2 = hlStr.ToUpper().Trim();
            if (s2 == "HOSPITAL" || s2 == "H")
            {
                return HealthLevel.Hospital;
            }
            if (s2 == "COMMUNITY" || s2 == "C")
            {
                return HealthLevel.Community;
            }
            if (s2 == "TOTAL" || s2 == "T")
            {
                return HealthLevel.Total;
            }
            throw new ArgumentException("Health level is not valid", hlStr);
        }

        public static string GetStringForHealthLevel(HealthLevel hl)
        {
            string value = "";
            switch (hl)
            {
                case HealthLevel.Hospital:
                    value = "H";
                    break;
                case HealthLevel.Community:
                    value = "C";
                    break;
                case HealthLevel.Total:
                    value = "T";
                    break;
                default:
                    break;
            }
            return value;
        }
    }  
}
