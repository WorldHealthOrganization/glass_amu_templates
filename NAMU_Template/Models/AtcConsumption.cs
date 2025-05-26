// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using AMU_Template.Constants;
using NAMU_Template.Constants;

namespace NAMU_Template.Models
{
    public class AtcConsumption
    {
        public string Country { get; set; }

        public int Year { get; set; }

        public HealthSector Sector { get; set; }

        public string Level { get; set; }

        public string AMClass { get; set; }
        public string AtcClass { get; set; }

        public string AWaRe { get; set; }

        public YesNoNA MEML { get; set; }

        public string ATC5 { get; set; }

        public string ATC4 { get; set; }

        public string ATC3 { get; set; }
        
        public string ATC2 { get; set; }

        public string Roa { get; set; }

        public YesNoUnknown Paediatric { get; set; }

        public bool AvailabilityTotal { get; set; }

        public bool AvailabilityHospital{ get; set; }

        public bool AvailabilityCommunity{ get; set; }

        public Decimal PKGConsumptionTotal { get; private set; }

        public Decimal PKGConsumptionCommunity { get; private set; }

        public Decimal PKGConsumptionHospital { get; private set; }

        public Decimal DIDConsumptionTotal { get; private set; }

        public Decimal DIDConsumptionCommunity { get; private set; }

        public Decimal DIDConsumptionHospital { get; private set; }

        public Decimal DDDConsumptionTotal { get; private set; }

        public Decimal DDDConsumptionCommunity { get; private set; }

        public Decimal DDDConsumptionHospital { get; private set; }

        
        public AtcConsumption()
        {
            PKGConsumptionTotal = 0;
            PKGConsumptionCommunity = 0;
            PKGConsumptionHospital = 0;

            DIDConsumptionTotal = 0;
            DIDConsumptionCommunity = 0;
            DIDConsumptionHospital = 0;

            DDDConsumptionTotal = 0;
            DDDConsumptionCommunity = 0;
            DDDConsumptionHospital = 0;
        }

        public bool IsConsumptionData()
        {
            return PKGConsumptionTotal > 0 || PKGConsumptionCommunity > 0 || PKGConsumptionHospital > 0;
        }

        public bool IsConsumptionCommunityData()
        {
            return PKGConsumptionTotal > 0;
        }

        public bool IsHospitalConsumptionData()
        {
            return PKGConsumptionHospital > 0;
        }

        public void AddProductConsumption(ProductConsumption prodCons)
        {
            PKGConsumptionTotal += prodCons.PKGConsumptionTotal;
            DDDConsumptionTotal += prodCons.DDDConsumptionTotal;
            DIDConsumptionTotal += prodCons.DIDConsumptionTotal;

            PKGConsumptionCommunity += prodCons.PKGConsumptionCommunity;
            DDDConsumptionCommunity += prodCons.DDDConsumptionCommunity;
            DIDConsumptionCommunity += prodCons.DIDConsumptionCommunity;

            PKGConsumptionHospital += prodCons.PKGConsumptionHospital;
            DDDConsumptionHospital += prodCons.DDDConsumptionHospital;
            DIDConsumptionHospital += prodCons.DIDConsumptionHospital;
        }

    }
}
