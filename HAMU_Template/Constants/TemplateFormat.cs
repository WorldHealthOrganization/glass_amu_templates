// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using System.Windows.Forms;

namespace HAMU_Template.Constants
{
    public static class TemplateFormat
    {
        // Worksheet names
        public static string LEGAL_SHEETNAME = "WHO Template";
        public static string MACRO_SHEETNAME = "Macro";
        public static string AVAILABILITY_SHEETNAME = "Availability Data";
        public static string HOSPITAL_STRUCTURE_SHEETNAME = "Hospital Structure";
        public static string HOSPITAL_ACTIVITY_SHEETNAME = "Hospital Activity Data";
        public static string PRODUCT_SHEETNAME = "Product Data";
        public static string SUBSTANCE_SHEETNAME = "Substance Data";
        public static string ATC_SHEETNAME = "ATC";
        public static string DDD_SHEETNAME = "DDD";
        public static string COMBINATION_SHEETNAME = "DDD Combination";
        public static string CONVERSION_SHEETNAME = "Conversion";
        public static string UNIT_SHEETNAME = "Units";
        public static string SALT_SHEETNAME = "Salts";
        public static string ROA_SHEETNAME = "RoAs";
        public static string MEML_SHEETNAME = "mEML";
        public static string AWARE_SHEETNAME = "AWaRe";
        public static string PRODUCT_ORIGIN_SHEETNAME = "Product Origins";

        public static string DATA_SHEETNAME = PRODUCT_SHEETNAME;

        public static int DATA_SHEET_STATUS_COL_IDX = 1;
        public static int DATA_SHEET_STATUS_MSG_COL_IDX = 2;
        public static int DATA_SHEET_AUTO_CAL_COL_NB = 9;


        // PRODUCT

        public static int PRODUCT_DATA_SHEET_START_DATA_ROW_IDX = 3;

        public static int PRODUCT_DATA_SHEET_STATUS_COL_IDX = 1;
        public static int PRODUCT_DATA_SHEET_STATUS_MSG_COL_IDX = 2;
        public static int PRODUCT_DATA_SHEET_COUNTRY_COL_IDX = 3;
        public static int PRODUCT_DATA_SHEET_HOSPITAL_COL_IDX = 4;
        public static int PRODUCT_DATA_SHEET_PRODUCT_ID_COL_IDX = 5;
        public static int PRODUCT_DATA_SHEET_LABEL_COL_IDX = 6;
        public static int PRODUCT_DATA_SHEET_PACKSIZE_COL_IDX = 7;
        public static int PRODUCT_DATA_SHEET_ROA_COL_IDX = 8;
        public static int PRODUCT_DATA_SHEET_STRENGTH_COL_IDX = 9;
        public static int PRODUCT_DATA_SHEET_STRENGTH_UNIT_COL_IDX = 10;
        public static int PRODUCT_DATA_SHEET_CONCENTRATION_COL_IDX = 11;
        public static int PRODUCT_DATA_SHEET_VOLUME_COL_IDX = 12;
        public static int PRODUCT_DATA_SHEET_ATC5_COL_IDX = 13;
        public static int PRODUCT_DATA_SHEET_SALT_COL_IDX = 14;
        public static int PRODUCT_DATA_SHEET_COMBINATION_COL_IDX = 15;
        public static int PRODUCT_DATA_SHEET_PAEDIATRIC_COL_IDX = 16;
        public static int PRODUCT_DATA_SHEET_FORM_COL_IDX = 17;
        public static int PRODUCT_DATA_SHEET_PRODUCT_NAME_COL_IDX = 18;
        public static int PRODUCT_DATA_SHEET_INGREDIENTS_COL_IDX = 19;
        public static int PRODUCT_DATA_SHEET_ORIGIN_COL_IDX = 20;
        public static int PRODUCT_DATA_SHEET_GENERICS_COL_IDX = 21;
        public static int PRODUCT_DATA_SHEET_CONVERSION_FACTOR_COL_IDX = 22;
        public static int PRODUCT_DATA_SHEET_CONTENT_COL_IDX = 23;
        public static int PRODUCT_DATA_SHEET_CONTENT_UNIT_COL_IDX = 24;
        public static int PRODUCT_DATA_SHEET_ARS_COL_IDX = 25;
        public static int PRODUCT_DATA_SHEET_DDD_VALUE_COL_IDX = 26;
        public static int PRODUCT_DATA_SHEET_DDD_UNIT_COL_IDX = 27;
        public static int PRODUCT_DATA_SHEET_AWR_COL_IDX = 28;
        public static int PRODUCT_DATA_SHEET_MEML_COL_IDX = 29;
        public static int PRODUCT_DATA_SHEET_DPP_COL_IDX = 30;


        // SUBSTANCE

        public static int SUBSTANCE_DATA_SHEET_START_DATA_ROW_IDX = 3;

        public static int SUBSTANCE_DATA_SHEET_STATUS_COL_IDX = 1;
        public static int SUBSTANCE_DATA_SHEET_STATUS_MSG_COL_IDX = 2;
        public static int SUBSTANCE_DATA_SHEET_COUNTRY_COL_IDX = 3;
        public static int SUBSTANCE_DATA_SHEET_HOSPITAL_COL_IDX = 4;
        public static int SUBSTANCE_DATA_SHEET_LABEL_COL_IDX = 5;
        public static int SUBSTANCE_DATA_SHEET_ROA_COL_IDX = 6;
        public static int SUBSTANCE_DATA_SHEET_STRENGTH_COL_IDX = 7;
        public static int SUBSTANCE_DATA_SHEET_STRENGTH_UNIT_COL_IDX = 8;
        public static int SUBSTANCE_DATA_SHEET_CONCENTRATION_COL_IDX = 9;
        public static int SUBSTANCE_DATA_SHEET_VOLUME_COL_IDX = 10;
        public static int SUBSTANCE_DATA_SHEET_ATC5_COL_IDX = 11;
        public static int SUBSTANCE_DATA_SHEET_SALT_COL_IDX = 12;
        public static int SUBSTANCE_DATA_SHEET_COMBINATION_COL_IDX = 13;
        public static int SUBSTANCE_DATA_SHEET_PAEDIATRIC_COL_IDX = 14;
        public static int SUBSTANCE_DATA_SHEET_FORM_COL_IDX = 15;
        public static int SUBSTANCE_DATA_SHEET_INGREDIENTS_COL_IDX = 16;
        public static int SUBSTANCE_DATA_SHEET_CONVERSION_FACTOR_COL_IDX = 17;
        public static int SUBSTANCE_DATA_SHEET_CONTENT_COL_IDX = 18;
        public static int SUBSTANCE_DATA_SHEET_CONTENT_UNIT_COL_IDX = 19;
        public static int SUBSTANCE_DATA_SHEET_ARS_COL_IDX = 20;
        public static int SUBSTANCE_DATA_SHEET_DDD_VALUE_COL_IDX = 21;
        public static int SUBSTANCE_DATA_SHEET_DDD_UNIT_COL_IDX = 22;
        public static int SUBSTANCE_DATA_SHEET_AWR_COL_IDX = 23;
        public static int SUBSTANCE_DATA_SHEET_MEML_COL_IDX = 24;
        public static int SUBSTANCE_DATA_SHEET_DPP_COL_IDX = 25;


        public static string CONS_START_COL_IDX = "CONS_START_COL_IDX ";
        public static string AUTO_CALC_START_COL_IDX = "AUTO_CALC_START_COL_IDX";
        public static string AUTO_CALC_END_COL_IDX = "AUTO_CALC_END_COL_ID";
        public static string AUTO_CALC_CONVERSION_FACTOR_COL_IDX = "AUTO_CALC_CONVERSION_FACTOR_COL_IDX";
        public static string AUTO_CALC_CONTENT_COL_IDX = "AUTO_CALC_CONTENT_COL_IDX";
        public static string AUTO_CALC_CONTENT_UNIT_COL_IDX = "AUTO_CALC_CONTENT_UNIT_COL_IDX";
        public static string AUTO_CALC_ARS_COL_IDX = "AUTO_CALC_ARS_COL_IDX";
        public static string AUTO_CALC_DDD_VALUE_COL_IDX = "AUTO_CALC_DDD_VALUE_COL_IDX";
        public static string AUTO_CALC_DDD_UNIT_COL_IDX = "AUTO_CALC_DDD_UNIT_COL_IDX";
        public static string AUTO_CALC_AWR_COL_IDX = "AUTO_CALC_AWR_COL_IDX";
        public static string AUTO_CALC_MEML_COL_IDX = "AUTO_CALC_MEML_COL_IDX";
        public static string AUTO_CALC_DPP_COL_IDX = "AUTO_CALC_DPP_COL_IDX";



        public static int SUBSTANCE_DATA_SHEET_CONS_START_COL_IDX = 26;
        public static int SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX = SUBSTANCE_DATA_SHEET_CONVERSION_FACTOR_COL_IDX;
        public static int SUBSTANCE_DATA_SHEET_AUTO_CALC_END_COL_IDX = SUBSTANCE_DATA_SHEET_DPP_COL_IDX;

        public static int PRODUCT_DATA_SHEET_CONS_START_COL_IDX = 31;
        public static int PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX = PRODUCT_DATA_SHEET_CONVERSION_FACTOR_COL_IDX;
        public static int PRODUCT_DATA_SHEET_AUTO_CALC_END_COL_IDX = PRODUCT_DATA_SHEET_DPP_COL_IDX;

        public static IDictionary<string, int> SUBSTANCE_COL_IDX_MAP = new Dictionary<string, int>()
        {
            { CONS_START_COL_IDX, SUBSTANCE_DATA_SHEET_CONS_START_COL_IDX },
            { AUTO_CALC_START_COL_IDX, SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_END_COL_IDX, SUBSTANCE_DATA_SHEET_AUTO_CALC_END_COL_IDX },
            { AUTO_CALC_CONVERSION_FACTOR_COL_IDX,  SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_CONTENT_COL_IDX, SUBSTANCE_DATA_SHEET_CONTENT_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_CONTENT_UNIT_COL_IDX, SUBSTANCE_DATA_SHEET_CONTENT_UNIT_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_ARS_COL_IDX, SUBSTANCE_DATA_SHEET_ARS_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_DDD_VALUE_COL_IDX, SUBSTANCE_DATA_SHEET_DDD_VALUE_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_DDD_UNIT_COL_IDX, SUBSTANCE_DATA_SHEET_DDD_UNIT_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_AWR_COL_IDX, SUBSTANCE_DATA_SHEET_AWR_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_MEML_COL_IDX, SUBSTANCE_DATA_SHEET_MEML_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_DPP_COL_IDX, SUBSTANCE_DATA_SHEET_DPP_COL_IDX - SUBSTANCE_DATA_SHEET_AUTO_CALC_START_COL_IDX },
        };

        public static IDictionary<string, int> PRODUCT_COL_IDX_MAP = new Dictionary<string, int>()
        {
            { CONS_START_COL_IDX, PRODUCT_DATA_SHEET_CONS_START_COL_IDX },
            { AUTO_CALC_START_COL_IDX, PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_END_COL_IDX, PRODUCT_DATA_SHEET_AUTO_CALC_END_COL_IDX },
            { AUTO_CALC_CONVERSION_FACTOR_COL_IDX,  PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_CONTENT_COL_IDX, PRODUCT_DATA_SHEET_CONTENT_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_CONTENT_UNIT_COL_IDX, PRODUCT_DATA_SHEET_CONTENT_UNIT_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_ARS_COL_IDX, PRODUCT_DATA_SHEET_ARS_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_DDD_VALUE_COL_IDX, PRODUCT_DATA_SHEET_DDD_VALUE_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_DDD_UNIT_COL_IDX, PRODUCT_DATA_SHEET_DDD_UNIT_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_AWR_COL_IDX, PRODUCT_DATA_SHEET_AWR_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_MEML_COL_IDX, PRODUCT_DATA_SHEET_MEML_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
            { AUTO_CALC_DPP_COL_IDX, PRODUCT_DATA_SHEET_DPP_COL_IDX - PRODUCT_DATA_SHEET_AUTO_CALC_START_COL_IDX },
        };


        // HOSPITAL STRUCTURE
        public static int STRUCTURE_SHEET_COUNTRY_COL_IDX = 1;
        public static int STRUCTURE_SHEET_HOSPITAL_COL_IDX = 2;
        public static int STRUCTURE_SHEET_YEAR_COL_IDX = 3;


        // AVAILABILITY
        public static int AVAILABILITY_SHEET_COUNTRY_COL_IDX = 1;
        public static int AVAILABILITY_SHEET_HOSPITAL_COL_IDX = 2;
        public static int AVAILABILITY_SHEET_YEAR_COL_IDX = 3;
        public static int AVAILABILITY_SHEET_LEVEL_COL_IDX = 4;
        public static int AVAILABILITY_SHEET_ATC_COL_IDX_START = 5;
        public static IDictionary<string, int> AVAILABILITY_SHEET_ATC_COL_IDX_MAP = new Dictionary<string, int>()
        {
            {"A07AA", 5},
            {"D01BA", 6},
            {"J01", 7},
            {"J02", 8},
            {"J04", 9},
            {"J05", 10},
            {"P01AB", 11},
            {"P01B", 12}
        };

        // HOSPITAL ACTIVITY
        public static int ACTIVITY_SHEET_COUNTRY_COL_IDX = 1;
        public static int ACTIVITY_SHEET_HOSPITAL_COL_IDX = 2;
        public static int ACTIVITY_SHEET_YEAR_COL_IDX = 3;
        public static int ACTIVITY_SHEET_STRUCTURE_COL_IDX = 4;
        public static int ACTIVITY_SHEET_PATIENT_DAYS_COL_IDX = 5;
        public static int ACTIVITY_SHEET_ADMISSIONS_COL_IDX = 6;
    }
}
