// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace HAMU_Template.Constants
{
    public enum FacilityStructureLevel
    {
        Hospital = 1,
        Department = 2,
        Ward = 3
    }

    public class FacilityStructureLevelString
    {
        public static FacilityStructureLevel GetFacilityStructureLevelForString(string slStr)
        {
            string s2 = slStr.ToUpper().Trim();
            if (s2 == "HOSPITAL" || s2 == "H")
            {
                return FacilityStructureLevel.Hospital;
            }
            if (s2 == "DEPARTMENT" || s2 == "D")
            {
                return FacilityStructureLevel.Department;
            }
            if (s2 == "WARD" || s2 == "W")
            {
                return FacilityStructureLevel.Ward;
            }
            throw new ArgumentException("Facility structure level is not valid", slStr);
        }

        public static string GetStringForFacilityStructureLevel(FacilityStructureLevel sl)
        {
            string value = "";
            switch (sl)
            {
                case FacilityStructureLevel.Hospital:
                    value = "HOSPITAL";
                    break;
                case FacilityStructureLevel.Department:
                    value = "DEPARTMENT";
                    break;
                case FacilityStructureLevel.Ward:
                    value = "WARD";
                    break;
                default:
                    break;
            }
            return value;
        }
    }
}