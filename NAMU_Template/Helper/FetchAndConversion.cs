// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace NAMU_Template.Helper
{
    public static class FetchAndConversion
    {
        /*public static List<CombinedDDD> listCombDdds = ThisWorkbook.CombinedDDD_DataList;
        public static List<DDDClass> listDDDs = ThisWorkbook.DDD_DataList;
        public static List<MeasureUnit> listUnits = ThisWorkbook.UnitDataList;

        public static CombinedDDD FetchWhoCombinedDdd(string comb)
        {
            CombinedDDD d = listCombDdds.FirstOrDefault(item => item.Code == comb);
            return d;
        }
        public static DDDClass FetchWhoDdd(string ars)
        {
            // Check if the ARS exists in the dictionary
            DDDClass d = listDDDs.FirstOrDefault(factor => factor.ARS == ars);
            return d;
        }
        public static MeasureUnit FetchUnit(string unitCode)
        {
            MeasureUnit u = listUnits.FirstOrDefault(item => item.Code == unitCode);
            return u;
        }
        public static double? ConvertBaseUnit(double amount, string unitMeasure)
        {
            // Fetch the unit from the list
            MeasureUnit u = listUnits.FirstOrDefault(item => item.Code == unitMeasure);

            if (u != null)
            {
                double? newAmount = amount * u.BaseConversion;
                return newAmount;
            }
            return null;
        }
*/
    }
}
