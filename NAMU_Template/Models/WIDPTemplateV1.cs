// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;

namespace NAMU_Template.Models
{
    public class WIDPTemplateV1
    {
        // Private fields
        private string _name;
        private string _version;

        private Dictionary<string, object[]> _worksheets;
        private Dictionary<string, Dictionary<string, WidpVariable>> _variables;

        // Constructor
        public WIDPTemplateV1()
        {
            Initialize();
        }

        // Initialization
        public void Initialize()
        {
            _name = "AMR_GLASS_AMC_WIDP_V1";
            _version = "20240410";

            _worksheets = new Dictionary<string, object[]>();
            _variables = new Dictionary<string, Dictionary<string, WidpVariable>>();

            PopulateWorksheets();
            PopulateVariables();
        }

        // Methods for worksheet properties
        public string GetRegisterWorksheetName()
        {
            return (string)_worksheets["REGISTER"][1];
        }

        public int GetRegisterWorksheetIndex()
        {
            return (int)_worksheets["REGISTER"][0];
        }

        public int GetRegisterWorksheetStartRow()
        {
            return (int)_worksheets["REGISTER"][2];
        }

        public string GetUseWorksheetName()
        {
            return (string)_worksheets["USE"][1];
        }

        public int GetUseWorksheetIndex()
        {
            return (int)_worksheets["USE"][0];
        }

        public int GetUseWorksheetStartRow()
        {
            return (int)_worksheets["USE"][2];
        }

        // Methods for retrieving column index
        public int GetColumnIndexForREGVariable(string varName)
        {
            if (_variables["REGISTER"].ContainsKey(varName))
            {
                return _variables["REGISTER"][varName].Index;
            }
            throw new KeyNotFoundException($"Variable '{varName}' not found in REGISTER.");
        }

        public int GetColumnIndexForUSEVariable(string varName)
        {
            if (_variables["USE"].ContainsKey(varName))
            {
                return _variables["USE"][varName].Index;
            }
            throw new KeyNotFoundException($"Variable '{varName}' not found in USE.");
        }

        // Populate worksheets
        private void PopulateWorksheets()
        {
            _worksheets.Add("REGISTER", new object[] { 1, "TEI Instances", 6 });
            _worksheets.Add("USE", new object[] { 2, "(1) AMC - Raw Product Consumpti", 3 });
        }

        // Populate variables
        private void PopulateVariables()
        {
            _variables.Add("REGISTER", new Dictionary<string, WidpVariable>());
            _variables.Add("USE", new Dictionary<string, WidpVariable>());

            PopulateRegisterVariables();
            PopulateUseVariables();
        }

        private void PopulateRegisterVariables()
        {
            var regVariables = _variables["REGISTER"];
            regVariables.Add(Product.UID_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TE_UID", "TEI id", 1));
            regVariables.Add(Product.COUNTRY_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TE_ORG_UNIT", "Org Unit", 2));
            regVariables.Add(Product.ENROLMENT_DATE_WIDP_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TE_ENROL_DATE", "Enrollment Date", 4));
            regVariables.Add(Product.INCIDENT_DATE_WIDP_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TE_INCID_DATE", "Incident Date", 5));
            regVariables.Add(Product.PRODUCT_ID_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_PRODUCT_ID", "Product id", 6));
            regVariables.Add(Product.PRODUCT_NAME_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_PRODUCT_NAME", "Product name", 7));
            regVariables.Add(Product.LABEL_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_LABEL", "Label", 8));
            regVariables.Add(Product.PACKSIZE_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_PACKSIZE", "Package size", 9));
            regVariables.Add(Product.STRENGTH_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_STRENGTH", "Strength", 10));
            regVariables.Add(Product.STRENGTH_UNIT_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_STRENGTH_UNIT", "Strength unit", 11));
            regVariables.Add(Product.CONCENTRATION_VOLUME_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_CONC_VOLUME", "Concentration volume", 12));
            regVariables.Add(Product.VOLUME_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_VOLUME", "Volume", 13));
            regVariables.Add(Product.ATC5_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_ATC", "ATC Code", 14));
            regVariables.Add(Product.COMBINATION_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_COMBINATION", "Combination Code", 15));
            regVariables.Add(Product.ROUTE_ADMIN_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_ROUTE_ADMIN", "Route of administration", 16));
            regVariables.Add(Product.SALT_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_SALT", "Salt", 17));
            regVariables.Add(Product.PAEDIATRIC_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_PAEDIATRIC_PRODUCT", "Is it a paediatric product?", 18));
            regVariables.Add(Product.FORM_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_FORM", "formName", 19));
            regVariables.Add(Product.INGREDIENTS_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_INGREDIENTS", "Ingredients", 20));
            regVariables.Add(Product.PRODUCT_ORIGIN_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_PRODUCT_ORIGIN", "origin", 21));
            regVariables.Add(Product.MANUFACTURER_COUNTRY_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_MANUFACTURER_COUNTRY", "Manufacturer country", 22));
            regVariables.Add(Product.MARKET_AUTH_HOLDER_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_MARKET_AUTH_HOLDER", "Marketing authorization holder", 23));
            regVariables.Add(Product.GENERICS_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_GENERICS", "Is it a generic product?", 24));
            regVariables.Add(Product.YEAR_AUTHORIZATION_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_YEAR_AUTHORIZATION", "Authorization year", 25));
            regVariables.Add(Product.YEAR_WITHDRAWAL_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_YEAR_WITHDRAWAL", "Withdrawal year", 26));
            regVariables.Add(Product.DATA_STATUS_WIDP_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TEA_DATA_STATUS", "Data status", 27));

        }

        private void PopulateUseVariables()
        {
            var useVariables = _variables["USE"];
            useVariables.Add(ProductConsumption.PROD_CONS_PROD_UID_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TE_TEI", "TEI id", 2));
            useVariables.Add(ProductConsumption.PROD_CONS_OPTIONS_WIDP_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TE_OPTIONS", "Options", 3));
            useVariables.Add(ProductConsumption.PROD_CONS_EVENT_DATE_WIDP_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_TE_EVENT_DATE", "Date", 4));
            useVariables.Add(ProductConsumption.PROD_CONS_H_SECTOR_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_DET_SECTOR_MANUAL", "health_sector_manual", 5));
            useVariables.Add(ProductConsumption.PROD_CONS_H_LEVEL_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_DET_LEVEL_MANUAL", "health_level_manual", 6));
            useVariables.Add(ProductConsumption.PROD_CONS_PACKAGES_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_DET_PACKAGES", "packages_det", 7));
            useVariables.Add(ProductConsumption.PROD_CONS_STATUS_WIDP_FIELD, CreateWIDPVariable("AMR_GLASS_AMC_DET_DATA_STATUS_MANUAL", "data_status_manual", 8));
            // Add remaining variables as needed...
        }

        // Factory method for creating variables
        private WidpVariable CreateWIDPVariable(string id, string name, int index)
        {
            return new WidpVariable(id, name, index);
        }
    }
}
