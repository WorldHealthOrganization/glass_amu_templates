// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using NAMU_Template.Helper;
using NAMU_Template.Models;

namespace NAMU_Template.Data_Export
{
    public static class ExportConsumption
    {
        public static void ExportCalculateUseConsumption(
            List<DataAvailability> availData)
        {
            // Retrieve the calculated DDD Consumption data from SharedData
            var atcConsumptionData = SharedData.AtcConsumptionData;
            var productConsumptionData = SharedData.ProductConsummptionData;
            var productData = SharedData.Products;
            // Create a new Excel application and a new workbook
            var excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            //Call methods to export ATC consumption
            ExportATC.ExportATCConsumption(atcConsumptionData, availData, workbook);

            //Call methods to export Product consumption
            ExportATC.ExportProductConsumption(productConsumptionData, productData, availData, workbook);

            excelApp.Visible = true;
        }


    }
}
