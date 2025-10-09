// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace AMU_Template.Constants
{
    public static class AMUConstants
    {

        public const string YES = "YES";
        public const string NO = "NO";
        public const string UNK = "UNK";
        public const string NA = "NA";
        public const string ATC_Z99_CODE = "Z99ZZ99";

        public const string COMB_Z99_CODE = "Z99ZZ99_99";

        public static AdministrationRoute ROA_X = new AdministrationRoute { Code = "X", Name = "UNDEFINED" };

        public static MeasureUnit UNIT_X = new MeasureUnit
        {
            Code = "X",
            Family = MeasureUnitFamily.Undefined,
            BaseConversion = new Decimal(1)
        };

        
        public static ATC ATC_Z99 = new ATC {
            Code = ATC_Z99_CODE,
            Name = "UNDEFINED ATC",
            Level = 5,
            AMClass = "NC",
            ATCClass = "Z99"
        };

        public static DDDCombination COMB_Z99 = new DDDCombination
        {
            Code = COMB_Z99_CODE,
            ATC5 = ATC_Z99,
            Form = "UNDEFINED",
            ROA = ROA_X,
            UnitDose = "UNDEFINED",
            DDDValue = 0,
            DDDUnit = UNIT_X
        };
            
       

        public static readonly IList<string> ATCClasses = new ReadOnlyCollection<string>
        (new List<String> {
        "A07AA", "D01BA", "J01", "J02", "J04", "J05", "P01AB", "P01B", "Z99" });

        public static readonly IDictionary<string, string> AMClasses = new ReadOnlyDictionary<string, string>
        (new Dictionary<string, string> {
            { "ATB","Antibiotics" }, {"ATF","Antifungals" }, {"ATV","Antivirals" }, {"ATM","Antimalarials" }, {"ATT", "Antituberculosis"}, {"OTH","Others"}});


        public static readonly IList<string> AWARECategories = new ReadOnlyCollection<string>(new List<string>
        {
            "A","W","R","N","N/A"
        });

        public static readonly IList<string> MEMLCategories = new ReadOnlyCollection<string>(new List<string>
        {
            "YES","NO","N/A"
        });
    }

    public enum YesNoUnknown
    {
        Unknown = -1,
        No = 1,
        Yes = 2
    }

    public enum YesNoNA
    {
        NA = -1,
        No = 1,
        Yes = 2
    }


    public static class YesNoUnknownString
    {
        public static YesNoUnknown GetYesNoUnkFromString(string val)
        {
            string val2 = val.ToUpper().Trim();
            if (val2 == "Y" || val2 == "YES")
            {
                return YesNoUnknown.Yes;
            }
            if (val2 == "N" || val2 == "NO")
            {
                return YesNoUnknown.No;
            }
            if (val2 == "U" || val2 == "UNK" || val2 == "UNKNOWN")
            {
                return YesNoUnknown.Unknown;
            }
            throw new ArgumentException($"YesNoUnknown value {val2} is not valid");
        }

        public static string GetStringFromYesNoUnk(YesNoUnknown ynu)
        {
            string val = "";
            switch (ynu)
            {
                case YesNoUnknown.Yes:
                    val = AMUConstants.YES; break;
                case YesNoUnknown.No:
                    val = AMUConstants.NO; break;
                case YesNoUnknown.Unknown:
                    val = AMUConstants.UNK; break;
                default:
                    break;
            }
            return val;
        }
    }

    public static class YesNoNAString
    {
        public static YesNoNA GetYesNoNAFromString(string val)
        {
            string val2 = val.ToUpper().Trim();
            if (val2 == "Y" || val2 == "YES")
            {
                return YesNoNA.Yes;
            }
            if (val2 == "N" || val2 == "NO")
            {
                return YesNoNA.No;
            }
            if (val2 == "NA" || val2 == "N/A")
            {
                return YesNoNA.NA;
            }
            throw new ArgumentException($"YesNoNA value {val2} is not valid");
        }

        public static string GetStringFromYesNoNA(YesNoNA ynna)
        {
            string val = "";
            switch (ynna)
            {
                case YesNoNA.Yes:
                    val = AMUConstants.YES; break;
                case YesNoNA.No:
                    val = AMUConstants.NO; break;
                case YesNoNA.NA:
                    val = AMUConstants.NA; break;
                default:
                    break;
            }
            return val;
        }
    }

    public enum VStatus
    {
        NA = 0,
        DIRTY = 1,
        PARSED = 2,
        CALCULATED = 3,
        EXPORTED = 4,
    }

    public class VStatusString
    {
        public static IDictionary<VStatus, string> VStatus2String = new ReadOnlyDictionary<VStatus, string>(new Dictionary<VStatus, string> { { VStatus.NA, "NA" }, { VStatus.DIRTY, "Modified" }, { VStatus.PARSED, "Parsed and validated" }, { VStatus.CALCULATED, "Calculated" }, { VStatus.EXPORTED, "Exported" } });
    }
}

