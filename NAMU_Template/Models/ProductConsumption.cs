// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using AMU_Template.Constants;
using NAMU_Template.Constants;

namespace NAMU_Template.Models
{
    public class ProductConsumption
    {

        // WIDP Excel Template Fields
        public const string PROD_CONS_PROD_UID_FIELD = "PROD_CONS_PROD_UID_FIELD";
        public const string PROD_CONS_OPTIONS_WIDP_FIELD = "PROD_CONS_OPTIONS_WIDP_FIELD";
        public const string PROD_CONS_EVENT_DATE_WIDP_FIELD = "PROD_CONS_EVENT_DATE_WIDP_FIELD";
        public const string PROD_CONS_H_SECTOR_FIELD = "PROD_CONS_H_SECTOR_FIELD";
        public const string PROD_CONS_H_LEVEL_FIELD = "PROD_CONS_H_LEVEL_FIELD";
        public const string PROD_CONS_PACKAGES_FIELD = "PROD_CONS_PACKAGES_FIELD";
        public const string PROD_CONS_STATUS_WIDP_FIELD = "PROD_CONS_STATUS_WIDP_FIELD";

        public string Key { get; set; }

        public int LineNo { get; set; }

        public int Sequence { get; set; }

        public string Country { get; set; }

        public int Year { get; set; }

        public HealthSector Sector { get; set; }

        public string Level { get; set; }

        public string AMClass { get; set; }
        public string AtcClass { get; set; }
        
        public string AWaRe { get; set; }

        public YesNoNA MEML { get; set; }

        public string ProductId { get; set; }

        public string ProductUniqueId { get; set; }

        public string Label { get; set; }

        public string ATC5 { get; set; }

        private string InternalATC4;

        public string ATC4 { get { 
                if (!string.IsNullOrEmpty(InternalATC4)) { return InternalATC4; }
                else
                {
                    if (!string.IsNullOrEmpty(ATC5)) {
                        InternalATC4 = ATC5.Substring(0,5);
                        return InternalATC4;
                    }
                    else
                    {
                        return null;
                    }
                }
            } }

        private string InternalATC3;
        public string ATC3
        {
            get
            {
                if (!string.IsNullOrEmpty(InternalATC3)) { return InternalATC3; }
                else
                {
                    if (!string.IsNullOrEmpty(ATC5))
                    {
                        InternalATC3 = ATC5.Substring(0, 4);
                        return InternalATC3;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }

        private string InternalATC2;
        public string ATC2
        {
            get
            {
                if (!string.IsNullOrEmpty(InternalATC2)) { return InternalATC2; }
                else
                {
                    if (!string.IsNullOrEmpty(ATC5))
                    {
                        InternalATC2 = ATC5.Substring(0, 3);
                        return InternalATC2;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }

        public string Roa { get; set; }

        public YesNoUnknown Paediatric { get; set; }  
       
        public Decimal DPP { get; set; }

        public Decimal PKGConsumptionTotal { get; set; }

        public Decimal PKGConsumptionCommunity { get; set; }

        public Decimal PKGConsumptionHospital { get; set; }

        public Decimal PopulationTotal { get; set; }

        public Decimal PopulationCommunity { get; set; }

        public Decimal PopulationHospital { get; set; }

        public Decimal DDDConsumptionTotal { get; set; }

        public Decimal DDDConsumptionCommunity { get; set; }

        public Decimal DDDConsumptionHospital { get; set; }

        public Decimal DIDConsumptionTotal { get; set; }

        public Decimal DIDConsumptionCommunity { get; set; }

        public Decimal DIDConsumptionHospital { get; set; }

        public bool AvailabilityTotal { get; set; }

        public bool AvailabilityCommunity { get; set; }

        public bool AvailabilityHospital { get; set; }

        public ProductConsumption()
        {
            PopulationTotal = 0;
            PopulationCommunity = 0;
            PopulationHospital = 0;

            DDDConsumptionTotal = 0;
            DDDConsumptionCommunity = 0;
            DDDConsumptionHospital = 0;

            DIDConsumptionTotal = 0;
            DIDConsumptionCommunity = 0;
            DIDConsumptionHospital = 0;

            PKGConsumptionTotal = 0;
            PKGConsumptionHospital = 0;
            PKGConsumptionCommunity = 0;

            AvailabilityTotal = false;
            AvailabilityCommunity = false;
            AvailabilityHospital = false;
        }

        public void CalculateDDD()
        {
            if (DPP > 0)
            {
                DDDConsumptionTotal = PKGConsumptionTotal * DPP;
                DDDConsumptionCommunity = PKGConsumptionCommunity * DPP;
                DDDConsumptionHospital = PKGConsumptionHospital * DPP;
            }
        }

        public void CalculateDID()
        {
            if (DPP > 0)
            {
                if (PopulationTotal != 0)
                {
                    DIDConsumptionTotal = DDDConsumptionTotal * 1000 / (365 * PopulationTotal);
                }
                if (PopulationCommunity != 0)
                {
                    DIDConsumptionCommunity = DDDConsumptionCommunity * 1000 / (365 * PopulationCommunity);
                }

                if (PopulationHospital != 0)
                {
                    DIDConsumptionHospital = DDDConsumptionHospital * 1000 / (365 * PopulationHospital);
                }
            }
        }

        public object GetValueForVariable(string variable)
        {
            object value;

            switch (variable)
            {
                case PROD_CONS_PROD_UID_FIELD:
                    value = this.ProductUniqueId; // what is equivalent  to uid?
                    break;

                case PROD_CONS_OPTIONS_WIDP_FIELD:
                    value = null;
                    break;

                case PROD_CONS_EVENT_DATE_WIDP_FIELD:
                    value = new DateTime((int)Year, 1, 1);
                    break;

                case PROD_CONS_H_SECTOR_FIELD:
                    value = Sector;
                    break;

                case PROD_CONS_H_LEVEL_FIELD:
                    value = Level;
                    break;

                case PROD_CONS_PACKAGES_FIELD:
                    value = Label;
                    break;

                case PROD_CONS_STATUS_WIDP_FIELD:
                    value = 1;
                    break;

                default:
                    value = "";
                    break;
            }

            return value;
        }

    }
}
