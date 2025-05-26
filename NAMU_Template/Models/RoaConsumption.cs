// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;

namespace NAMU_Template.Models
{
    public class RoaConsumption
    {
        public string Roa { get; set; }

        public AtcConsumption Parent { get; set; }

        public Decimal PKGTotal { get; set; }

        public Decimal PKGCommunity { get; set; }

        public Decimal PKGHospital { get; set; }

        public Decimal DIDTotal { get; set; }

        public Decimal DIDCommunity { get; set; }

        public Decimal DIDHospital { get; set; }

        public Decimal DDDTotal { get; set; }

        public Decimal DDDCommunity { get; set; }

        public Decimal DDDHospital { get; set; }

        public RoaConsumption()
        {
            DIDTotal = 0;
            DIDCommunity = 0;
            DIDHospital = 0;

            DDDTotal = 0;
            DDDHospital = 0;
            DDDCommunity = 0;

            PKGTotal = 0;
            PKGCommunity = 0;
            PKGHospital = 0;
        }

        public bool IsConsumptionData()
        {
            return PKGTotal > 0 || PKGCommunity > 0 || PKGHospital > 0;
        }

        public bool IsTotalConsumptionData()
        {
            return PKGTotal > 0;
        }

        public bool IsCommunityConsumptionData()
        {
            return PKGCommunity > 0;
        }

        public bool IsHospitalConsumptionData()
        {
            return PKGHospital > 0;
        }

        public void AddProductConsumption(ProductConsumption prodCons)
        {
            PKGTotal += prodCons.PKGConsumptionTotal;
            DDDTotal += prodCons.DDDConsumptionTotal;
            DIDTotal += prodCons.DIDConsumptionTotal;

            PKGCommunity = PKGCommunity + prodCons.PKGConsumptionCommunity;
            DDDCommunity = DDDCommunity + prodCons.DDDConsumptionCommunity;
            DIDCommunity = DIDCommunity + prodCons.DIDConsumptionCommunity;

            PKGHospital = PKGHospital + prodCons.PKGConsumptionHospital;
            DDDHospital = DDDHospital + prodCons.DDDConsumptionHospital;
            DIDHospital = DIDHospital + prodCons.DIDConsumptionHospital;


        }
    }
}
