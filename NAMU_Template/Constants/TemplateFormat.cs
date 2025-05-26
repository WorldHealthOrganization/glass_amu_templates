// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace NAMU_Template.Constants
{
    public static class TemplateFormat
    {
        // Worksheet names
        public static string LEGAL_SHEETNAME = "WHO Template";
        public static string MACRO_SHEETNAME = "Macro";
        public static string AVAILABILITY_SHEETNAME = "Availability Data";
        public static string POPULATION_SHEETNAME = "Population Data";
        public static string PRODUCT_SHEETNAME = "Product Data";
        public static string ATC_SHEETNAME = "ATC";
        public static string DDD_SHEETNAME = "DDD";
        public static string COMBINATION_SHEETNAME = "DDD Combination";
        public static string CONVERSION_SHEETNAME = "Conversion";
        public static string UNIT_SHEETNAME = "Units";
        public static string SALT_SHEETNAME = "Salts";
        public static string ROA_SHEETNAME = "RoAs";
        public static string MEML_SHEETNAME = "mEML";
        public static string AWARE_SHEETNAME = "AWaRe";
        public static string COUNTRY_SHEETNAME = "Countries";
        public static string PRODUCT_ORIGIN_SHEETNAME = "Product Origins";

        public static string DATA_SHEETNAME = PRODUCT_SHEETNAME;


        // PRODUCT
        public static int DATA_SHEET_START_DATA_ROW_IDX = 3;
        public static int DATA_SHEET_STATUS_COL_IDX = 1;
        public static int DATA_SHEET_STATUS_MSG_COL_IDX = 2;
        public static int DATA_SHEET_COUNTRY_COL_IDX = 3;
        public static int DATA_SHEET_PRODUCT_ID_COL_IDX = 4;
        public static int DATA_SHEET_LABEL_COL_IDX = 5;
        public static int DATA_SHEET_PACKSIZE_COL_IDX = 6;
        public static int DATA_SHEET_ROA_COL_IDX = 7;
        public static int DATA_SHEET_STRENGTH_COL_IDX = 8;
        public static int DATA_SHEET_STRENGTH_UNIT_COL_IDX = 9;
        public static int DATA_SHEET_CONCENTRATION_COL_IDX = 10;
        public static int DATA_SHEET_VOLUME_COL_IDX = 11;
        public static int DATA_SHEET_ATC5_COL_IDX = 12;
        public static int DATA_SHEET_SALT_COL_IDX = 13;
        public static int DATA_SHEET_COMBINATION_COL_IDX = 14;
        public static int DATA_SHEET_PAEDIATRIC_COL_IDX = 15;
        public static int DATA_SHEET_FORM_COL_IDX = 16;
        public static int DATA_SHEET_PRODUCT_NAME_COL_IDX = 17;
        public static int DATA_SHEET_INGREDIENTS_COL_IDX = 18;
        public static int DATA_SHEET_ORIGIN_COL_IDX = 19;
        public static int DATA_SHEET_MAN_COUNTRY_COL_IDX = 20;
        public static int DATA_SHEET_MAH_COL_IDX = 21;
        public static int DATA_SHEET_GENERICS_COL_IDX = 22;
        public static int DATA_SHEET_YEAR_AUTHORIZATION_COL_IDX = 23;
        public static int DATA_SHEET_YEAR_WITHDRAWAL_COL_IDX = 24;
        public static int DATA_SHEET_CONVERSION_FACTOR_COL_IDX = 25;
        public static int DATA_SHEET_CONTENT_COL_IDX = 26;
        public static int DATA_SHEET_CONTENT_UNIT_COL_IDX = 27;
        public static int DATA_SHEET_ARS_COL_IDX = 28;
        public static int DATA_SHEET_DDD_VALUE_COL_IDX = 29;
        public static int DATA_SHEET_DDD_UNIT_COL_IDX = 30;
        public static int DATA_SHEET_AWR_COL_IDX = 31;
        public static int DATA_SHEET_MEML_COL_IDX = 32;
        public static int DATA_SHEET_DPP_COL_IDX = 33;
        public static int DATA_SHEET_SECTOR_COL_IDX = 34;

        public static int DATA_SHEET_PACKS_TOTAL_COL_IDX = 0;
        public static int DATA_SHEET_PACKS_COMMUNITY_COL_IDX = 1;
        public static int DATA_SHEET_PACKS_HOSPITAL_COL_IDX = 2;

        public static int DATA_SHEET_PACKS_START_COL_IDX = 35;
        public static int DATA_SHEET_PACKS_LENGTH_COLS = 3;

        public static int AUTO_CALC_START_COL_IDX = DATA_SHEET_CONVERSION_FACTOR_COL_IDX;

        public static int AUTO_CALC_CONVERSION_FACTOR_COL_IDX = DATA_SHEET_CONVERSION_FACTOR_COL_IDX - AUTO_CALC_START_COL_IDX;
        public static int AUTO_CALC_CONTENT_COL_IDX = DATA_SHEET_CONTENT_COL_IDX - AUTO_CALC_START_COL_IDX;
        public static int AUTO_CALC_CONTENT_UNIT_COL_IDX = DATA_SHEET_CONTENT_UNIT_COL_IDX - AUTO_CALC_START_COL_IDX;
        public static int AUTO_CALC_ARS_COL_IDX = DATA_SHEET_ARS_COL_IDX - AUTO_CALC_START_COL_IDX;
        public static int AUTO_CALC_DDD_VALUE_COL_IDX = DATA_SHEET_DDD_VALUE_COL_IDX - AUTO_CALC_START_COL_IDX;
        public static int AUTO_CALC_DDD_UNIT_COL_IDX = DATA_SHEET_DDD_UNIT_COL_IDX - AUTO_CALC_START_COL_IDX;
        public static int AUTO_CALC_AWR_COL_IDX = DATA_SHEET_AWR_COL_IDX - AUTO_CALC_START_COL_IDX;
        public static int AUTO_CALC_MEML_COL_IDX = DATA_SHEET_MEML_COL_IDX - AUTO_CALC_START_COL_IDX;
        public static int AUTO_CALC_DPP_COL_IDX = DATA_SHEET_DPP_COL_IDX - AUTO_CALC_START_COL_IDX;

        public static int AUTO_CALC_END_COL_IDX = AUTO_CALC_DPP_COL_IDX;


        // POPULATION
        public static int POP_SHEET_COUNTRY_COL_IDX = 1;
        public static int POP_SHEET_YEAR_COL_IDX = 2;
        public static int POP_SHEET_ATCCLASS_COL_IDX = 3;
        public static int POP_SHEET_SECTOR_COL_IDX = 4;
        public static int POP_SHEET_POP_TOTAL_COL_IDX = 5;
        public static int POP_SHEET_POP_COMMUNITY_COL_IDX = 6;
        public static int POP_SHEET_POP_HOSPITAL_COL_IDX = 7;

        // AVAILABILITY
        public static int AVAILABILITY_START_YEAR_COL_INDEX = 5;
    }
}
