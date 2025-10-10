// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using HAMU_Template.Models;
using System.Collections.Generic;
using System.Linq;

namespace HAMU_Template.Data_Processing
{
    public static class ConsumptionProcessing
    {

        public static List<AtcConsumption> CalculateAtcDDDConsumption(List<MedicineConsumption> medConsData)
        {
            // Validate that the input dictionary is not null or empty
            if (medConsData == null || medConsData.Count == 0)
            {
                return null; // Exit the function as there's nothing to process
            }

            // Create a dictionary of year->amClass->atc->sector->productId->DataConsumption
            // Create a list to store DataConsumption objects
            var atcConsumptionData = new List<AtcConsumption>();

            foreach (var medCons in medConsData)
            {
                if (medCons.ATC5 == "Z99ZZ99")
                {
                    continue;
                }

                var cntry = medCons.Country;
                var year = medCons.Year;
                var hospital = medCons.Hospital;
                var level = medCons.Level;
                var structure = medCons.Structure;
                var amClass = medCons.AMClass;
                var atcClass = medCons.AtcClass;
                var aware = medCons.AWaRe;
                var meml = medCons.MEML;
                var atc5 = medCons.ATC5;
                var atc4 = medCons.ATC4;
                var atc3 = medCons.ATC3;
                var atc2 = medCons.ATC2;
                var roa = medCons.Roa;
                var paed = medCons.Paediatric;

                // Create or find the SubstanceConsumption for this combination of values
                var existingAtcCons = atcConsumptionData
                    .FirstOrDefault(d => d.Country == cntry && d.Year == year && d.Hospital == hospital &&
                        d.AtcClass == atcClass && d.ATC5 == atc5 && d.Roa == roa && d.Paediatric == paed);
                if (existingAtcCons == null)
                {
                    // If it does not exist, create a new one
                    existingAtcCons = new AtcConsumption
                    {
                        Country = cntry,
                        Year = year,
                        Hospital = hospital,
                        Level = level,
                        Structure = structure,
                        AMClass = amClass,
                        AtcClass = atcClass,
                        AWaRe = aware, // Add AWaRe classification
                        MEML = meml,    // Add mEML status
                        ATC5 = atc5,
                        ATC4 = atc4,
                        ATC3 = atc3,
                        ATC2 = atc2,
                        Roa = roa,
                        Paediatric = paed,
                    };

                    // Add the new DataConsumption to the list
                    atcConsumptionData.Add(existingAtcCons);
                }

                // Add product consumption to the existing DataConsumption
                existingAtcCons.AddMedicineConsumption(medCons);
            }

            // Store the result in sharedData
            return atcConsumptionData;
        }
    }
}
