// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Numerics;
using AMU_Template.Constants;
using HAMU_Template.Constants;

namespace HAMU_Template.Models
{
    public class AtcConsumption
    {
        public string Country { get; set; }

        public string Hospital { get; set; }

        public int Year { get; set; }

        public FacilityStructureLevel Level { get; set; }

        public string Structure { get; set; }

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

        public Decimal PKGUnitConsumption { get; set; }

        public Decimal DDDConsumption { get; set; }

        public Decimal DADConsumption { get; private set; }

        public Decimal DBDConsumption { get; private set; }
        
        public AtcConsumption()
        {
            PKGUnitConsumption = 0;
            DDDConsumption = 0;
            DADConsumption = 0;
            DBDConsumption = 0;
        }

        public bool IsConsumptionData()
        {
            return PKGUnitConsumption > 0;
        }

        public void AddMedicineConsumption(MedicineConsumption medCons)
        {
            PKGUnitConsumption += medCons.GetPackUnit();
            DDDConsumption += medCons.DDD;
            DADConsumption += medCons.DAD;
            DBDConsumption += medCons.DBD;
        }
    }
}
