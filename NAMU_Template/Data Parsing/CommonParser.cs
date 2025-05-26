// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using AMU_Template.Constants;
using NAMU_Template.Constants;
using NAMU_Template.Data_Validation;

namespace NAMU_Template.Data_Parsing
{
    public static class CommonParser
    {
        public static string ParseString(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (cellValue == null || String.IsNullOrEmpty(cellValue.ToString()))
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory";
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                    es.AddErrorMsgs(errMsg);
                }
                return string.Empty;
            }
            return cellValue.ToString().Trim();
        }

        public static bool? ParseBoolean(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {

            if (cellValue == null)
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory.";
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                    es.AddErrorMsgs(errMsg);
                }

                return null;
            }

            string val = cellValue.ToString().ToUpper().Trim();

            if (val.Equals("Y") || val.Equals("YES"))
            {
                return true;
            }
            else if (val.Equals("N") || val.Equals("NO"))
            {
                return false;
            }
            else
            {
                string errMsg = $"{variable} has invalid value {val} for a boolean. It should be Y or N";
                variableErrors += errMsg + '\n';
                es.Status = EntityStatus.ERROR;
                es.AddErrorMsgs(errMsg);
                return null;
            }
        }

        public static Decimal? ParseDecimal(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (cellValue == null || cellValue.ToString().Trim() == "")
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory.";
                    es.AddErrorMsgs(errMsg);
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                }
                return null;
            }

            if (Decimal.TryParse(cellValue.ToString(), out Decimal result))
            {
                return result; // Convert double to int with rounding
            }
            else
            {
                string errMsg = $"{variable} has an invalid decimal value.";
                es.AddErrorMsgs(errMsg);
                variableErrors += errMsg + "\n";
                es.Status = EntityStatus.ERROR;
                return null;
            }
        }

        public static int? ParseInteger(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (cellValue == null || cellValue.ToString().Trim() == "")
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory.";
                    es.AddErrorMsgs(errMsg);
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                }
                return null;
            }

            if (int.TryParse(cellValue.ToString(), out int result))
            {
                return result; 
            }
            else
            {
                string errMsg = $"{variable} has an invalid value.";
                es.AddErrorMsgs(errMsg);
                variableErrors += errMsg + "\n";
                es.Status = EntityStatus.ERROR;
                return null;
            }
        }

        public static double? ParseNumber(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (cellValue == null || cellValue.ToString().Trim() == "")
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory.";
                    es.AddErrorMsgs(errMsg);
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                }
                return null;
            }

            if (double.TryParse(cellValue.ToString(), out double result))
            {
                return result; // Convert double to int with rounding
            }
            else
            {
                string errMsg = $"{variable} has an invalid value for a number.";
                es.AddErrorMsgs(errMsg);
                variableErrors += errMsg + "\n";
                es.Status = EntityStatus.ERROR;
                return null;
            }
        }

        public static string ParseCountryISO3(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (cellValue == null || String.IsNullOrEmpty(cellValue.ToString()))
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory";
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                    es.AddErrorMsgs(errMsg);
                }
                return string.Empty;
            }
            string val = cellValue.ToString().Trim().ToUpper();
            if (val.Length != 3)
            {
                string errMsg = $"{variable} is not a valid ISO3 country code (3 letters)";
                variableErrors += errMsg + '\n';
                es.Status = EntityStatus.ERROR;
                es.AddErrorMsgs(errMsg);
                
                return string.Empty;
            }
            return val;
        }

        public static int? ParseYear(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (cellValue == null || String.IsNullOrEmpty(cellValue.ToString()))
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory.";
                    es.AddErrorMsgs(errMsg);
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                }
                return null;
            }

            if (int.TryParse(cellValue.ToString(), out int result))
            {
                return result;
            }
            else
            {
                string errMsg = $"{variable} has an invalid value {cellValue.ToString()} for decimal.";
                es.AddErrorMsgs(errMsg);
                variableErrors += errMsg + "\n";
                es.Status = EntityStatus.ERROR;
                return null;
            }
        }

        public static string ParseATCClass(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (cellValue == null || String.IsNullOrEmpty(cellValue.ToString()))
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory";
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                    es.AddErrorMsgs(errMsg);
                }
                return string.Empty;
            }
            string val = cellValue.ToString().Trim().ToUpper();
            
            if (!AMUConstants.ATCClasses.Contains(val))
            {
                string errMsg = $"{variable} is not a valid ATC class";
                variableErrors += errMsg + '\n';
                es.Status = EntityStatus.ERROR;
                es.AddErrorMsgs(errMsg);

                return string.Empty;
            }
            return val;
        }

        public static HealthSector? ParseHealthSector(dynamic cellValue, string variable, ErrorStatus es, ref string variableErrors, bool mandatory)
        {
            if (cellValue == null || String.IsNullOrEmpty(cellValue.ToString()))
            {
                if (mandatory)
                {
                    string errMsg = $"{variable} is mandatory";
                    variableErrors += errMsg + '\n';
                    es.Status = EntityStatus.ERROR;
                    es.AddErrorMsgs(errMsg);
                }
                return null;
            }
            string val = cellValue.ToString().Trim().ToUpper();

            try
            {
                HealthSector hs = HealthSectorLevelString.GetHealthSectorForString(val);
                return hs;
            }
            catch (Exception ex)
            {
                string errMsg = $"{variable} is not a valid health sector";
                variableErrors += errMsg + '\n';
                es.Status = EntityStatus.ERROR;
                es.AddErrorMsgs(errMsg);

                return null;
            }
        }


    }
}
