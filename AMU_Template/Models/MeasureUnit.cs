// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;

namespace AMU_Template.Models
{

    public enum MeasureUnitFamily
    {
        Gram,
        InternationalUnit,
        DoseUnit,
        Undefined
    }

    public static class MeasureUnitFamilyString
    {
        public static MeasureUnitFamily GetMeasureUnitFamilyFromString(string ufs)
        {
            MeasureUnitFamily uf;

            switch (ufs) {
                case "GRAM":
                    uf = MeasureUnitFamily.Gram;
                    break;
                case "INTERNATIONAL_UNIT":
                    uf = MeasureUnitFamily.InternationalUnit;
                    break;
                case "UNIT_DOSE":
                    uf = MeasureUnitFamily.DoseUnit;
                    break;
                default:
                    throw new Exception($"Invalid measurement unit family {ufs}.");
            }
            return uf;
        }
    }

    public class MeasureUnit
    {
        public string Code { get; set; }

        public MeasureUnitFamily Family { get; set; }

        public Decimal BaseConversion { get; set; }

        public MeasureUnit()
        {

        }

        public MeasureUnit(string code, MeasureUnitFamily family, Decimal baseConversion) 
        { 
            Code = code;
            Family = family;
            BaseConversion = baseConversion;
        }

        public Decimal convertValueToBaseUnit(Decimal value)
        {
            return value*BaseConversion;
        }
    }
}
