// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using AMU_Template.Models;
using NAMU_Template.Constants;
using NAMU_Template.Data_Parsing;
using NAMU_Template.Models;


namespace NAMU_Template
{
    public partial class ThisWorkbook
    {

        public static Dictionary<string, ATC> ATCDataDict;
        public static Dictionary<string, ATC> ATC5DataDict;
        public static Dictionary<string, DDD> DDDDataDict;
        public static Dictionary<string, DDDCombination> DDDCombinationDataDict;
        public static List<ConversionFactor> ConversionFactorDataList;
        public static Dictionary<string, MeasureUnit> UnitDataDict;
        public static Dictionary<string, Salt> SaltDataDict;
        public static Dictionary<string, AdministrationRoute> AdminRouteDataDict;
        public static Dictionary<string, Country> CountryDataDict;
        public static Dictionary<string, ProductOrigin> ProductOriginDataDict;
        public static List<Aware> AwareDataList;
        public static List<MEML> MemlDataList;
        public static VStatus VSTATUS;

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {

            Application.ActiveWindow.Zoom = 100;

            VSTATUS = Constants.VStatus.NA;

            //Not to send the whole workbook in every function..!
            Excel.Worksheet ATC_Datasheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.ATC_SHEETNAME];
            Excel.Worksheet DDD_Datasheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.DDD_SHEETNAME];
            Excel.Worksheet CombinedDDD_Datasheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.COMBINATION_SHEETNAME];
            Excel.Worksheet ConversionFactor_Datasheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.CONVERSION_SHEETNAME];
            Excel.Worksheet Unit_Datasheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.UNIT_SHEETNAME];
            Excel.Worksheet Salt_Datasheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.SALT_SHEETNAME];
            Excel.Worksheet AdminRoute_Datasheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.ROA_SHEETNAME];
            Excel.Worksheet Aware_DataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.AWARE_SHEETNAME];
            Excel.Worksheet mEML_DataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.MEML_SHEETNAME];
            Excel.Worksheet Country_DataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.COUNTRY_SHEETNAME];
            Excel.Worksheet ProductOrigin_DataSheet = this.Application.ActiveWorkbook.Worksheets[TemplateFormat.PRODUCT_ORIGIN_SHEETNAME];

            var atcList = ReferenceDataParser.ProcessATC(ATC_Datasheet);
            ATCDataDict = atcList.ToDictionary(a => a.Code);

            var atc5List = ATCDataDict.Values.Where(a => a.Level == 5);
            ATC5DataDict = atc5List.ToDictionary(a => a.Code);

            var saltList = ReferenceDataParser.ProcessSalt(Salt_Datasheet);
            SaltDataDict = saltList.ToDictionary(s => s.Code);

            var roaList = ReferenceDataParser.ProcessRoAs(AdminRoute_Datasheet);
            AdminRouteDataDict = roaList.ToDictionary(r => r.Code);

            var unitList = ReferenceDataParser.ProcessUnit(Unit_Datasheet);
            UnitDataDict = unitList.ToDictionary(u => u.Code);

            var cntryList = ReferenceDataParser.ProcessCountry(Country_DataSheet);
            CountryDataDict = cntryList.ToDictionary(c => c.Code);

            var pOrigList = ReferenceDataParser.ProcessProductOrigin(ProductOrigin_DataSheet);
            ProductOriginDataDict = pOrigList.ToDictionary(o => o.Code);

            AwareDataList = ReferenceDataParser.ProcessAware(Aware_DataSheet);
            MemlDataList = ReferenceDataParser.ProcessMeml(mEML_DataSheet);

            var dddList = ReferenceDataParser.ProcessDDD(DDD_Datasheet, ATCDataDict, AdminRouteDataDict, SaltDataDict, UnitDataDict);
            DDDDataDict = dddList.ToDictionary(d => d.ARS);
            var combList = ReferenceDataParser.ProcessDDDCombination(CombinedDDD_Datasheet, ATCDataDict, AdminRouteDataDict, UnitDataDict);
            DDDCombinationDataDict = combList.ToDictionary(c => c.Code);
            ConversionFactorDataList = ReferenceDataParser.ProcessConversionFactor(ConversionFactor_Datasheet, ATCDataDict, AdminRouteDataDict, SaltDataDict, UnitDataDict);

            //Lock the sheets..!

            /*
            ATC_Datasheet.Protect("GLASS@2025");
            DDD_Datasheet.Protect("GLASS@2025");
            CombinedDDD_Datasheet.Protect("GLASS@2025");
            ConversionFactor_Datasheet.Protect("GLASS@2025");
            Unit_Datasheet.Protect("GLASS@2025");
            Salt_Datasheet.Protect("GLASS@2025");
            AdminRoute_Datasheet.Protect("GLASS@2025");
            Aware_DataSheet.Protect("GLASS@2025");
            mEML_DataSheet.Protect("GLASS@2025");
            */
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
