// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Constants;
using Excel = Microsoft.Office.Interop.Excel;

namespace NAMU_Template.Data_Validation
{
    public static class Validator
    {

        public static VStatus GetStatus()
        {
            return ThisWorkbook.VSTATUS;
        }

        public static void SetStatus(VStatus sts)
        {
            Feuil1 sheet = Globals.Feuil1;
            //if (VSTATUS != sts)
            //{
            //    VSTATUS = sts;
            //    Excel.Range statusCell = sheet.StatusCell;
            //    statusCell.Value2 = ThisWorkbook.CURRENT_STATUS_STRS[sts];
            //}

            if (ThisWorkbook.VSTATUS != sts)
            {
                ThisWorkbook.VSTATUS = sts;

                string statusValue = VStatusString.VStatus2String[ThisWorkbook.VSTATUS];

                //switch (sts)
                //{
                //    case Constants.VSTATUS_NA:
                //        statusValue = "NA";
                //        break;
                //    case Constants.VSTATUS_DIRTY:
                //        statusValue = "Modified";
                //        break;
                //    case Constants.VSTATUS_PARSED:
                //        statusValue = "Parsed and Validated";
                //        break;
                //    case Constants.VSTATUS_CALC:
                //        statusValue = "Calculated Consumption";
                //        break;
                //    case Constants.VSTATUS_EXPORT:
                //        statusValue = "Export Calculated Consumption";
                //        break;
                //    default:
                //        statusValue = "Unknown";
                //        break;
                //}

                // Update status in excel sheet
                Excel.Range statusCell = sheet.StatusCell;
                statusCell.Value2 = statusValue;

            }
        }
    }
}
