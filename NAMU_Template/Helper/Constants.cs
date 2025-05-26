// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace NAMU_Template.Helper
{
    public static class Constants
    {
        public const int HEADER1_ROW_IDX = 1;
        public const int HEADER2_ROW_IDX = 2;

        public const int PROD_ID_IDX = 1;
        public const string PROD_ID_NAME = "PRODUCT_ID";
        public const int PROD_LABEL_IDX = 2;
        public const string PROD_LABEL_NAME = "LABEL";
        public const int PROD_ATC_IDX = 3;
        public const string PROD_ATC_NAME = "ATC5";
        public const int PROD_ROA_IDX = 4;
        public const string PROD_ROA_NAME = "ROUTE_ADMIN";
        public const int DATA_START_IDX = 5;

        // Row refereces
        public const int START_DATA_ROW = 3; //main


        // Sheets references
        public const string MACRO_SHEET = "Macro";
        public const string AVAIL_SHEET = "Availability Data";
        public const string DATA_SHEET = "Product Data";
        public const string DATA_DDD_SHEET = "DDD Data";
        public const string POPULATION_SHEET = "Population Data";

        // Constanst for parseAvailability for loading Availability Data sheet
        public const string HOSPITAL_LEVEL = "H";
        public const string COMMUNITY_LEVEL = "C";
        public const string TOTAL_LEVEL = "T";
        public const int StartYearIndex = 5;

        //Status for validation and calculation

        public const int VSTATUS_NA = 0;
        public const string VSTATUS_NA_STR = "NA";
        public const int VSTATUS_DIRTY = 1;
        public const string VSTATUS_DIRTY_STR = "Modified";
        public const int VSTATUS_PARSED = 2;
        public const string VSTATUS_PARSED_STR = "Parsed and validated";
        public const int VSTATUS_CALC = 3;
        public const string VSTATUS_CALC_STR = "Use calculated";
        public const int VSTATUS_EXPORT = 4;


        private static int _vstatus = VSTATUS_NA;
        public static int VSTATUS
        {
            get => _vstatus;
            set
            {
                if (_vstatus != value)
                {
                    _vstatus = value;
                    //Console.WriteLine($"Status updated to: {_vstatus}");
                }
            }
        }

        // Cells reference for validation and calculation status
        public const string VSTATUS_IDX = "D4";
        public const string CSTATUS_IDX = "D5";

        //ATC Class references
        public const string A07AA_CLASS = "A07AA";
        public const string D01BA_CLASS = "D01BA";
        public const string J01_CLASS = "J01";
        public const string J02_CLASS = "J02";
        public const string J04_CLASS = "J04";
        public const string J05_CLASS = "J05";
        public const string P01AB_CLASS = "P01AB";
        public const string P01B_CLASS = "P01B";
        public const string D01BAJ02_CLASS = "D01BA_J02";
        public const string N04BB_CLASS = "N04BB";
        public const string UNK_AM_CLASS = "Z99";

        // Names of health sectors
        public const string PUBLIC_SECTOR = "PUB";
        public const string PRIVATE_SECTOR = "PRI";
        public const string GLOBAL_SECTOR = "GLO";

        // WIDP Constants
        public const string VAR_PROD_UID = "VAR_PROD_UID";
        public const string VAR_PROD_COUNTRY = "VAR_PROD_COUNTRY";
        public const string VAR_PROD_ENROL_DATE = "VAR_PROD_ENROL_DATE";
        public const string VAR_PROD_INCID_DATE = "VAR_PROD_INCID_DATE";
        public const string VAR_PROD_ID = "VAR_PROD_ID";
        public const string VAR_PROD_NAME = "VAR_PROD_NAME";
        public const string VAR_PROD_LABEL = "VAR_PROD_LABEL";
        public const string VAR_PROD_PACKSIZE = "VAR_PROD_PACKSIZE";
        public const string VAR_PROD_STRENGTH = "VAR_PROD_STRENGTH";
        public const string VAR_PROD_STRENGTH_UNIT = "VAR_PROD_STRENGTH_UNIT";
        public const string VAR_PROD_CONC_VOLUME = "VAR_PROD_CONC_VOLUME";
        public const string VAR_PROD_CONC_VOLUME_UNIT = "VAR_PROD_CONC_VOLUME_UNIT";
        public const string VAR_PROD_VOLUME = "VAR_PROD_VOLUME";
        public const string VAR_PROD_VOLUME_UNIT = "VAR_PROD_VOLUME_UNIT";
        public const string VAR_PROD_ATC = "VAR_PROD_ATC";
        public const string VAR_PROD_COMBINATION = "VAR_PROD_COMBINATION";
        public const string VAR_PROD_ROUTE_ADMIN = "VAR_PROD_ROUTE_ADMIN";
        public const string VAR_PROD_SALT = "VAR_PROD_SALT";
        public const string VAR_PROD_PAEDIATRIC_PRODUCT = "VAR_PROD_PAEDIATRIC_PRODUCT";
        public const string VAR_PROD_FORM = "VAR_PROD_FORM";
        public const string VAR_PROD_INGREDIENTS = "VAR_PROD_INGREDIENTS";
        public const string VAR_PROD_PRODUCT_ORIGIN = "VAR_PROD_PRODUCT_ORIGIN";
        public const string VAR_PROD_MANUFACTURER_COUNTRY = "VAR_PROD_MANUFACTURER_COUNTRY";
        public const string VAR_PROD_MARKET_AUTH_HOLDER = "VAR_PROD_MARKET_AUTH_HOLDER";
        public const string VAR_PROD_GENERICS = "VAR_PROD_GENERICS";
        public const string VAR_PROD_YEAR_AUTHORIZATION = "VAR_PROD_YEAR_AUTHORIZATION";
        public const string VAR_PROD_YEAR_WITHDRAWAL = "VAR_PROD_YEAR_WITHDRAWAL";
        public const string VAR_PROD_DATA_STATUS = "VAR_PROD_DATA_STATUS";

        public const string VALUE_UNIT_G = "G";
        public const string VALUE_UNIT_MG = "MG";
        public const string VALUE_UNIT_UD = "UD";
        public const string VALUE_UNIT_IU = "IU";
        public const string VALUE_UNIT_MIU = "MIU";

        public const int VALUE_DATA_STATUS = 1;


        public const string VALUE_ROA_O = "O";
        public const string VALUE_ROA_P = "P";
        public const string VALUE_ROA_IS = "IS";
        public const string VALUE_ROA_IP = "IP";
        public const string VALUE_ROA_R = "R";

        public const string VALUE_YNU_YES = "YES";
        public const string VALUE_YNU_NO = "NO";
        public const string VALUE_YNU_UNKNOWN = "UNK";

        public const string VALUE_SALT_DEFAULT = "XXXX";
        public const string VALUE_SALT_ESUC = "ESUC";
        public const string VALUE_SALT_HIPP = "HIPP";
        public const string VALUE_SALT_MAND = "MAND";
        public const string VALUE_PROD_ORIG_IMPORTED = "IMP";
        public const string VALUE_PROD_ORIG_PRODUCTION = "DOMPROD";
        public const string VALUE_PROD_ORIG_NGO = "NGO";
        public const string VALUE_PROD_ORIG_INTPROG = "INTPROG";

        public const string VAR_AMC_PROD_UID = "VAR_AMC_PROD_UID";
        public const string VAR_AMC_OPTIONS = "VAR_AMC_OPTIONS";
        public const string VAR_AMC_PACKS = "VAR_AMC_PACKS";
        public const string VAR_AMC_DATA_STATUS = "VAR_AMC_DATA_STATUS";
        public const string VAR_AMC_H_SECTOR = "VAR_AMC_H_SECTOR";
        public const string VAR_AMC_H_LEVEL = "VAR_AMC_H_LEVEL";
        public const string VAR_AMC_EVENT_DATE = "VAR_AMC_EVENT_DATE";

        public const string VALUE_H_SECTOR_PUBLIC = "PUB";
        public const string VALUE_H_SECTOR_PRIVATE = "PRI";
        public const string VALUE_H_SECTOR_GLOBAL = "GLO";
        public const string VALUE_H_LEVEL_COMMUNITY = "C";
        public const string VALUE_H_LEVEL_HOSPITAL = "H";
        public const string VALUE_H_LEVEL_TOTAL = "T";
    }
}
