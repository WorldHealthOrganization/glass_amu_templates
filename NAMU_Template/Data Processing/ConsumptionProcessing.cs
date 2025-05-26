// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using System.Linq;
using NAMU_Template.Helper;
using NAMU_Template.Models;
using static NAMU_Template.Helper.Constants;

namespace NAMU_Template.Data_Processing
{
    public static class ConsumptionProcessing
    {
        public static void ComputeConsumptionData(Dictionary<string, AtcConsumption> dataCons, Dictionary<string, DataAvailability> dataAvail)
        {
            /*
            var classes = new[]
            {
                A07AA_CLASS, D01BA_CLASS, J01_CLASS, J02_CLASS, J04_CLASS, J05_CLASS, P01AB_CLASS, N04BB_CLASS
            };

            
            foreach (var atcClass in classes)
            {
                if (dataAvail.TryGetValue(atcClass, out var avail) && avail.IsAvailable)
                {
                    switch (atcClass)
                    {
                        //case A07AA_CLASS:
                        //    ComputeA07AAConsumptionData(dataCons, avail);
                        //    break;
                        //case D01BA_CLASS:
                        //    ComputeD01BAConsumptionData(dataCons, avail);
                        //    break;
                        //case J01_CLASS:
                        //    ComputeJ01ConsumptionData(dataCons, avail);
                        //    break;
                        //case J02_CLASS:
                        //    ComputeJ02ConsumptionData(dataCons, avail);
                        //    break;
                        //case J04_CLASS:
                        //    ComputeJ04ConsumptionData(dataCons, avail);
                        //    break;
                        //case J05_CLASS:
                        //    ComputeJ05ConsumptionData(dataCons, avail);
                        //    break;
                        //case P01AB_CLASS:
                        //    ComputeP01ABConsumptionData(dataCons, avail);
                        //    break;
                        case N04BB_CLASS:
                            ComputeN04BBConsumptionData(dataCons, avail);
                            break;
                    }
                }
            }
            */
        }

        public static void ComputeJ01ConsumptionData(Dictionary<string, Dictionary<string, AtcConsumption>> dataCons,
                                                    DataAvailability dataAvail,
                                                    List<Product> listProducts)
        {
            string pattern = "J01";
            int patternLength = 3;
            double DIDY, PIY;

            // Ensure the "J01_CLASS" dictionary exists in `dataCons`
            if (!dataCons.ContainsKey(J01_CLASS))
            {
                dataCons["J01_CLASS"] = new Dictionary<string, AtcConsumption>();
            }

            foreach (var pr in listProducts)
            {
                // Check if the ATC class of the product starts with the specified pattern
                if (pr.ATC5.Code.StartsWith(pattern))
                {
                    // Ensure the product's ATC class exists in the dictionary
                    if (!dataCons["J01_CLASS"].ContainsKey(pr.ATC5.Code))
                    {
                        dataCons["J01_CLASS"][pr.ATC5.Code] = new AtcConsumption();
                    }

                    var data = dataCons["J01_CLASS"][pr.ATC5.Code];
                    //data.AtcClass = pr.ATC5;

                    // Add TOTAL availability data
                    if (dataAvail.AvailabilityTotal)
                    {
                        //DIDY = pr.DIDYearTotal;
                        //PIY = pr.PIYTotal;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("TOTAL", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("TOTAL", PIY);
                        //}
                    }

                    // Add COMMUNITY availability data
                    if (dataAvail.AvailabilityCommunity)
                    {
                        //DIDY = pr.DIDYearCommunity;
                        //PIY = pr.PIYCommunity;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("COMMUNITY", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("COMMUNITY", PIY);
                        //}
                    }

                    // Add HOSPITAL availability data
                    if (dataAvail.AvailabilityHospital)
                    {
                        //DIDY = pr.DIDYearHospital;
                        //PIY = pr.PIYHospital;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("HOSPITAL", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("HOSPITAL", PIY);
                        //}
                    }
                }
            }
        }

        public static void ComputeJ02ConsumptionData(
                                Dictionary<string, Dictionary<string, AtcConsumption>> dataCons,
                                DataAvailability dataAvail,
                                List<Product> listProducts)
        {
            string pattern = "J02";
            int patternLength = 3;
            double DIDY, PIY;

            // Ensure the "D01BAJ02_CLASS" dictionary exists in `dataCons`
            if (!dataCons.ContainsKey("D01BAJ02_CLASS"))
            {
                dataCons["D01BAJ02_CLASS"] = new Dictionary<string, AtcConsumption>();
            }

            foreach (var pr in listProducts)
            {
                // Check if the ATC class of the product starts with the specified pattern
                if (pr.ATC5.Code.StartsWith(pattern))
                {
                    // Ensure the product's ATC class exists in the dictionary
                    if (!dataCons["D01BAJ02_CLASS"].ContainsKey(pr.ATC5.Code))
                    {
                        dataCons["D01BAJ02_CLASS"][pr.ATC5.Code] = new AtcConsumption();
                    }

                    var data = dataCons["D01BAJ02_CLASS"][pr.ATC5.Code];
                    //data.AtcClass = pr.ATC5;

                    // Add TOTAL availability data
                    if (dataAvail.AvailabilityTotal)
                    {
                        //DIDY = pr.DIDYearTotal;
                        //PIY = pr.PIYTotal;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("TOTAL", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("TOTAL", PIY);
                        //}
                    }

                    // Add COMMUNITY availability data
                    if (dataAvail.AvailabilityCommunity)
                    {
                        //DIDY = pr.DIDYearCommunity;
                        //PIY = pr.PIYCommunity;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("COMMUNITY", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("COMMUNITY", PIY);
                        //}
                    }

                    // Add HOSPITAL availability data
                    if (dataAvail.AvailabilityHospital)
                    {
                        //DIDY = pr.DIDYearHospital;
                        //PIY = pr.PIYHospital;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("HOSPITAL", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("HOSPITAL", PIY);
                        //}
                    }
                }
            }
        }

        public static void ComputeJ04ConsumptionData(Dictionary<string, Dictionary<string, AtcConsumption>> dataCons,
                                                     DataAvailability dataAvail,
                                                     List<Product> listProducts)
        {
            string pattern = "J04";
            int patternLength = 3;
            double DIDY, PIY;

            // Ensure the "J04_CLASS" dictionary exists in `dataCons`
            if (!dataCons.ContainsKey("J04_CLASS"))
            {
                dataCons["J04_CLASS"] = new Dictionary<string, AtcConsumption>();
            }

            foreach (var pr in listProducts)
            {
                // Check if the ATC class of the product starts with the specified pattern
                if (pr.ATC5.Code.StartsWith(pattern))
                {
                    // Ensure the product's ATC class exists in the dictionary
                    if (!dataCons["J04_CLASS"].ContainsKey(pr.ATC5.Code))
                    {
                        dataCons["J04_CLASS"][pr.ATC5.Code] = new AtcConsumption();
                    }

                    var data = dataCons["J04_CLASS"][pr.ATC5.Code];
                    //data.AtcClass = pr.ATC5;

                    // Add TOTAL availability data
                    if (dataAvail.AvailabilityTotal)
                    {
                        //DIDY = pr.DIDYearTotal;
                        //PIY = pr.PIYTotal;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("TOTAL", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("TOTAL", PIY);
                        //}
                    }

                    // Add COMMUNITY availability data
                    if (dataAvail.AvailabilityCommunity)
                    {
                        //DIDY = pr.DIDYearCommunity;
                        //PIY = pr.PIYCommunity;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("COMMUNITY", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("COMMUNITY", PIY);
                        //}
                    }

                    // Add HOSPITAL availability data
                    if (dataAvail.AvailabilityHospital)
                    {
                        //DIDY = pr.DIDYearHospital;
                        //PIY = pr.PIYHospital;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("HOSPITAL", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("HOSPITAL", PIY);
                        //}
                    }
                }
            }
        }

        public static void ComputeD01BAConsumptionData(Dictionary<string, Dictionary<string, AtcConsumption>> dataCons,
                                                DataAvailability dataAvail,
                                                List<Product> listProducts)
        {
            string pattern = "D01BA";
            int patternLength = 5;
            double DIDY, PIY;

            // Ensure the "D01BAJ02_class"dictionary exists in dataCons
            if (!dataCons.ContainsKey(D01BAJ02_CLASS))
            {
                dataCons["D01BAJ02_CLASS"] = new Dictionary<string, AtcConsumption>();
            }

            foreach (var pr in listProducts)
            {
                // Check if the ATC class of the product starts with the specified pattern
                if (pr.ATC5.Code.StartsWith(pattern))
                {
                    //Ensure the product atc class exist in the dictionary 
                    if (!dataCons[D01BAJ02_CLASS].ContainsKey(pr.ATC5.Code))
                    {
                        dataCons["D01BAJ02_CLASS"][pr.ATC5.Code] = new AtcConsumption();
                    }

                    var data = dataCons["D01BAJ02_CLASS"][pr.ATC5.Code];
                    //data.AtcClass = pr.ATC5;


                    // Add total availability data 
                    if (dataAvail.AvailabilityTotal)
                    {
                        //DIDY = pr.DIDYearTotal;
                        //PIY = pr.PIYTotal;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("TOTAL", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("TOTAL", PIY);
                        //}
                    }

                    // Add community availability data
                    if (dataAvail.AvailabilityCommunity)
                    {
                        //DIDY = pr.DIDYearCommunity;
                        //PIY = pr.PIYCommunity;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("COMMUNITY", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("COMMUNITY", PIY);
                        //}
                    }

                    // Add HOSPITAL availability data
                    if (dataAvail.AvailabilityHospital)
                    {
                        //DIDY = pr.DIDYearHospital;
                        //PIY = pr.PIYHospital;

                        //if (double.IsFinite(DIDY))
                        //{
                        //    data.AddDIDConsumption("HOSPITAL", DIDY);
                        //}

                        //if (double.IsFinite(PIY))
                        //{
                        //    data.AddPIYConsumption("HOSPITAL", PIY);
                        //}
                    }
                }
            }
        }



        public static void ComputeN04BBConsumptionData(Dictionary<string, Dictionary<string, AtcConsumption>> dataCons, DataAvailability dataAvail)
        {
            const string pattern = "N04BB";
            const int patternLength = 5;

            if (!dataCons.ContainsKey(N04BB_CLASS))
                dataCons[N04BB_CLASS] = new Dictionary<string, AtcConsumption>();

            //foreach (var pr in ListProducts)
            //{
            //    if (pr.Atc5.StartsWith(pattern))
            //    {
            //        if (!dataCons[N04BB_CLASS].ContainsKey(pr.Atc5))
            //            dataCons[N04BB_CLASS][pr.Atc5] = new DataConsumption();

            //        var data = dataCons[N04BB_CLASS][pr.Atc5];
            //        data.AtcClass = pr.Atc5;

            //        if (dataAvail.AvailabilityTotal)
            //        {
            //            double didy = pr.DIDYearTotal;
            //            double piy = pr.PIYTotal;

            //            if (double.TryParse(didy.ToString(), out _))
            //                data.AddDIDConsumption("TOTAL", didy);
            //            if (double.TryParse(piy.ToString(), out _))
            //                data.AddPIYConsumption("TOTAL", piy);
            //        }

            //        if (dataAvail.AvailabilityCommunity)
            //        {
            //            double didy = pr.DIDYearCommunity;
            //            double piy = pr.PIYCommunity;

            //            if (double.TryParse(didy.ToString(), out _))
            //                data.AddDIDConsumption("COMMUNITY", didy);
            //            if (double.TryParse(piy.ToString(), out _))
            //                data.AddPIYConsumption("COMMUNITY", piy);
            //        }

            //        if (dataAvail.AvailabilityHospital)
            //        {
            //            double didy = pr.DIDYearHospital;
            //            double piy = pr.PIYHospital;

            //            if (double.TryParse(didy.ToString(), out _))
            //                data.AddDIDConsumption("HOSPITAL", didy);
            //            if (double.TryParse(piy.ToString(), out _))
            //                data.AddPIYConsumption("HOSPITAL", piy);
            //        }
            //    }
            // }
        }

        public static void ComputeA07AAConsumptionData(Dictionary<string, Dictionary<string, AtcConsumption>> dataCons, DataAvailability dataAvail)
        {
            const string pattern = "A07AA";
            const int patternLength = 5;

            if (!dataCons.ContainsKey(A07AA_CLASS))
            {
                dataCons[A07AA_CLASS] = new Dictionary<string, AtcConsumption>();
            }

            //foreach (var pr in listProducts)
            //{
            //    if (pr.Atc5.StartsWith(pattern))
            //    {
            //        if (!dataCons[A07AA_CLASS].ContainsKey(pr.Atc5))
            //        {
            //            dataCons[A07AA_CLASS][pr.Atc5] = new DataConsumption();
            //        }

            //        var data = dataCons[A07AA_CLASS][pr.Atc5];
            //        data.ATC = pr.Atc5;

            //        if (dataAvail.AvailabilityTotal)
            //        {
            //            double didy = pr.DIDYearTotal;
            //            double piy = pr.PIYTotal;

            //            if (double.TryParse(didy.ToString(), out _))
            //            {
            //                data.AddDIDConsumption("TOTAL", didy);
            //            }
            //            if (double.TryParse(piy.ToString(), out _))
            //            {
            //                data.AddPIYConsumption("TOTAL", piy);
            //            }
            //        }

            //        if (dataAvail.AvailabilityCommunity)
            //        {
            //            double didy = pr.DIDYearCommunity;
            //            double piy = pr.PIYCommunity;

            //            if (double.TryParse(didy.ToString(), out _))
            //            {
            //                data.AddDIDConsumption("COMMUNITY", didy);
            //            }
            //            if (double.TryParse(piy.ToString(), out _))
            //            {
            //                data.AddPIYConsumption("COMMUNITY", piy);
            //            }
            //        }

            //        if (dataAvail.AvailabilityHospital)
            //        {
            //            double didy = pr.DIDYearHospital;
            //            double piy = pr.PIYHospital;

            //            if (double.TryParse(didy.ToString(), out _))
            //            {
            //                data.AddDIDConsumption("HOSPITAL", didy);
            //            }
            //            if (double.TryParse(piy.ToString(), out _))
            //            {
            //                data.AddPIYConsumption("HOSPITAL", piy);
            //            }
            //        }
            //    }
            //}
        }

        public static void ComputeP01ABConsumptionData(Dictionary<string, Dictionary<string, AtcConsumption>> dataCons, DataAvailability dataAvail)
        {
            const string pattern = "P01AB";
            const int patternLength = 5;

            // Ensure the key exists in the dictionary
            if (!dataCons.ContainsKey(P01AB_CLASS))
            {
                dataCons[P01AB_CLASS] = new Dictionary<string, AtcConsumption>();
            }

            // Iterate through products
            //foreach (var pr in listProducts)
            //{
            //    if (pr.Atc5.StartsWith(pattern))
            //    {
            //        // Ensure the product key exists in the inner dictionary
            //        if (!dataCons[P01AB_CLASS].ContainsKey(pr.Atc5))
            //        {
            //            dataCons[P01AB_CLASS][pr.Atc5] = new DataConsumption();
            //        }

            //        var data = dataCons[P01AB_CLASS][pr.Atc5];
            //        data.ATC = pr.Atc5;

            //        // Process Total availability
            //        if (dataAvail.AvailabilityTotal)
            //        {
            //            double didy = pr.DIDYearTotal;
            //            double piy = pr.PIYTotal;

            //            if (double.TryParse(didy.ToString(), out _))
            //            {
            //                data.AddDIDConsumption("TOTAL", didy);
            //            }
            //            if (double.TryParse(piy.ToString(), out _))
            //            {
            //                data.AddPIYConsumption("TOTAL", piy);
            //            }
            //        }

            //        // Process Community availability
            //        if (dataAvail.AvailabilityCommunity)
            //        {
            //            double didy = pr.DIDYearCommunity;
            //            double piy = pr.PIYCommunity;

            //            if (double.TryParse(didy.ToString(), out _))
            //            {
            //                data.AddDIDConsumption("COMMUNITY", didy);
            //            }
            //            if (double.TryParse(piy.ToString(), out _))
            //            {
            //                data.AddPIYConsumption("COMMUNITY", piy);
            //            }
            //        }

            //        // Process Hospital availability
            //        if (dataAvail.AvailabilityHospital)
            //        {
            //            double didy = pr.DIDYearHospital;
            //            double piy = pr.PIYHospital;

            //            if (double.TryParse(didy.ToString(), out _))
            //            {
            //                data.AddDIDConsumption("HOSPITAL", didy);
            //            }
            //            if (double.TryParse(piy.ToString(), out _))
            //            {
            //                data.AddPIYConsumption("HOSPITAL", piy);
            //            }
            //        }
            //    }
            //}
        }


        public static void CalculateDDDConsumption(List<ProductConsumption> prodConsData, int[] years)
        {
            // Validate that the input dictionary is not null or empty
            if (prodConsData == null || prodConsData.Count == 0)
            {
                return; // Exit the function as there's nothing to process
            }

            // Create a dictionary of year->amClass->atc->sector->productId->DataConsumption
            // Create a list to store DataConsumption objects
            var atcConsumptionData = new List<AtcConsumption>();

            foreach (var prodCons in prodConsData)
            {
                if (prodCons.ATC5 == "Z99ZZ99")
                {
                    continue;
                }


                var cntry = prodCons.Country;
                var year = prodCons.Year;
                var sector = prodCons.Sector;
                var amClass = prodCons.AMClass;
                var atcClass = prodCons.AtcClass;
                var aware = prodCons.AWaRe;
                var meml = prodCons.MEML;
                var atc5 = prodCons.ATC5;
                var atc4 = prodCons.ATC4;
                var atc3 = prodCons.ATC3;
                var atc2 = prodCons.ATC2;
                var roa = prodCons.Roa;
                var paed = prodCons.Paediatric;
                var availT = prodCons.AvailabilityTotal;
                var availH = prodCons.AvailabilityHospital;
                var availC = prodCons.AvailabilityCommunity;
        
                // Create or find the SubstanceConsumption for this combination of values
                var existingAtcCons = atcConsumptionData
                    .FirstOrDefault(d => d.Country==cntry && d.Year == year && d.Sector == sector &&  
                        d.AMClass == amClass && d.AtcClass == atcClass && d.AWaRe == aware && d.MEML == meml && 
                         d.ATC5 == atc5 && d.Roa == roa && d.Paediatric == paed );
                if (existingAtcCons == null)
                {
                    // If it does not exist, create a new one
                    existingAtcCons = new AtcConsumption
                    {
                        Country = cntry,
                        Year = year,
                        Sector = sector,
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
                        AvailabilityCommunity = availC,
                        AvailabilityHospital = availH, 
                        AvailabilityTotal = availT
                    };

                    // Add the new DataConsumption to the list
                    atcConsumptionData.Add(existingAtcCons);
                }

                // Add product consumption to the existing DataConsumption
                existingAtcCons.AddProductConsumption(prodCons);
            }

            // Store the result in sharedData
            SharedData.AtcConsumptionData = atcConsumptionData;

            //PopulateProductConsumption(prodConsData);

            //SharedData.ProductConsummption = prodConsData;
        }


        public static void PopulateProductConsumption(List<ProductConsumption> prodConsData)
        {
            // Ensure the input data is valid
            if (prodConsData == null || prodConsData.Count == 0)
            {
                return; // Exit if there's nothing to process
            }
            var productConsumptionList = new List<ProductConsumption>();

            foreach (var prodCons in prodConsData)
            {
                // Find an existing product entry in the list
                var existingConsumption = productConsumptionList.FirstOrDefault(pc =>
                    pc.ProductId == prodCons.ProductId && pc.Year == prodCons.Year);

                if (existingConsumption == null)
                {
                    // Create a new ProductConsumption object and add it to the list
                    var newConsumption = new ProductConsumption
                    {
                        ProductId = prodCons.ProductId,
                        Year = prodCons.Year,
                        LineNo = prodCons.LineNo,
                        Sequence = prodCons.Sequence,
                        AMClass = prodCons.AMClass,
                        ATC5 = prodCons.ATC5,
                        Roa = prodCons.Roa,
                        Sector = prodCons.Sector,
                        AWaRe = prodCons.AWaRe,
                        MEML = prodCons.MEML,
                        AvailabilityTotal = prodCons.AvailabilityTotal,
                        AvailabilityCommunity = prodCons.AvailabilityCommunity,
                        AvailabilityHospital = prodCons.AvailabilityHospital,
                        PopulationTotal = prodCons.PopulationTotal,
                        PopulationCommunity = prodCons.PopulationCommunity,
                        PopulationHospital = prodCons.PopulationHospital,
                        DPP = prodCons.DPP,

                        PKGConsumptionTotal = prodCons.PKGConsumptionTotal,
                        PKGConsumptionCommunity = prodCons.PKGConsumptionCommunity,
                        PKGConsumptionHospital = prodCons.PKGConsumptionHospital,

                        DDDConsumptionTotal = prodCons.DDDConsumptionTotal,
                        DDDConsumptionCommunity = prodCons.DDDConsumptionCommunity,
                        DDDConsumptionHospital = prodCons.DDDConsumptionHospital,

                        DIDConsumptionTotal = prodCons.DIDConsumptionTotal,
                        DIDConsumptionCommunity = prodCons.DIDConsumptionCommunity,
                        DIDConsumptionHospital = prodCons.DIDConsumptionHospital,
                    };

                    productConsumptionList.Add(newConsumption);
                }
                else
                {
                    // Aggregate values if an existing record is found
                    existingConsumption.PKGConsumptionTotal += prodCons.PKGConsumptionTotal;
                    existingConsumption.DDDConsumptionTotal += prodCons.DDDConsumptionTotal;
                    existingConsumption.DIDConsumptionTotal += prodCons.DIDConsumptionTotal;

                    existingConsumption.PKGConsumptionCommunity += prodCons.PKGConsumptionCommunity;
                    existingConsumption.DDDConsumptionCommunity += prodCons.DDDConsumptionCommunity;
                    existingConsumption.DIDConsumptionCommunity += prodCons.DIDConsumptionCommunity;

                    existingConsumption.PKGConsumptionHospital += prodCons.PKGConsumptionHospital;
                    existingConsumption.DDDConsumptionHospital += prodCons.DDDConsumptionHospital;
                    existingConsumption.DIDConsumptionHospital += prodCons.DIDConsumptionHospital;
                }
                // Assign the processed list to SharedData
                SharedData.ProductConsummptionData = productConsumptionList;

            }
        }
    }
}
