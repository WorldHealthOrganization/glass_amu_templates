// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using AMU_Template.Models;

namespace NAMU_Template.Data_Parsing
{
    public class ProductValidator
    {

        public static List<Aware> listAware = ThisWorkbook.AwareDataList;
        public static List<MEML> listmEML = ThisWorkbook.MemlDataList;
        public static Dictionary<string, ATC> dictATC = ThisWorkbook.ATCDataDict;
        public static List<ConversionFactor> listConvFactors = ThisWorkbook.ConversionFactorDataList;
        public static Dictionary<string, MeasureUnit> listUnits = ThisWorkbook.UnitDataDict;

        /*public void ValidateProduct(Product pr, bool force)
        {
            if (pr.Infos.Count > 0)
            {
                pr.Status = EntityStatus.INFO;
            }

            if (pr.Warnings.Count > 0)
            {
                pr.Status = EntityStatus.WARNING;
            }

            if (pr.Errors.Count > 0)
            {
                pr.Status = EntityStatus.ERROR;
            }

            if (pr.Status > EntityStatus.OK && !force)
            {
                return;
            }


            ValidateAtc5(pr);
            ValidateRoA(pr);
            ValidateSalt(pr);
            CalculateArs(pr);

            ValidateCombination(pr);

            CalculateAware(pr);
            CalculateMEML(pr);

            ValidateProductId(pr);
            ValidateLabel(pr);
            ValidatePackageSize(pr);
            ValidateStrength(pr);
            ValidateConcentration(pr);
            ValidateDdds(pr);

            CalculateConvFactor(pr);
            CalculatePackageContent(pr);
            CalculateDddPerPackage(pr);

            ValidateYearsAuthorizationWithdrawal(pr);
            ValidatePaediatrics(pr);
            ValidateSector(pr);

        }

        private void ValidateAtc5(Product pr)
        {
            bool isAtc5ObjectValid = pr.ATC5 != null;
            pr.SetValidate(isAtc5ObjectValid, "ATC5");

            if (!isAtc5ObjectValid)
            {
                pr.AddWarningMsg("If there is no ATC for this product, use the code Z99ZZ99.");
            }
            else
            {
                if (pr.ATC5.Code == "Z99ZZ99")
                {
                    if (string.IsNullOrEmpty(pr.Ingredients)) // Need to verify
                    {
                        pr.SetValidate(false, "ATC5");
                        pr.AddErrorMsgs("As the ATC5 is Z99ZZ99, please provide the list of ingredients separated by comma in the ingredients column.");
                    }
                    pr.AddInfoMsg("Ensure that the ingredients and their respective strength are stated in the label.");
                }
            }
            
        }

        private void ValidateRoA(Product pr)
        {
            pr.SetValidate(pr.AdministrationRouteObject != null, "ROA");
        }

        private void ValidateSalt(Product pr)
        {
            pr.SetValidate(pr.SaltObject != null, "SALT");
        }

        private void ValidateCombination(Product pr)
        {
            var combs = new List<CombinedDDD>();

            if (pr.GetValidate("ARS") == false)
            {
                return;
            }


            pr.SetValidate(true, "COMBINATION");

            if (pr.CombinationObject != null && pr.CombinationObject.Code == "Z99ZZ99_99")
            {
                //if (pr.ATC5 != "Z99ZZ99")
                //{
                if (string.IsNullOrEmpty(pr.Ingredients)) // need to verify
                {
                    pr.AddErrorMsgs("The Combination code is set to undefined/Z99ZZ99_99. Please provide the list of ingredients separated by comma in the ingredients column.");
                    pr.SetValidate(false, "COMBINATION");
                }
                else
                {
                    pr.SetValidate(true, "COMBINATION");
                    pr.AddInfoMsg("The Combination code is set to undefined/Z99ZZ99_99.");
                }
                pr.AddInfoMsg("Ensure that the ingredients and their respective strength are stated in the label.");
                return;
                //}
            }

            if (pr.ATC5 == "Z99ZZ99" && pr.CombinationObject != null && pr.CombinationObject.Code != null)
            {
                pr.SetValidate(false, "COMBINATION");
                pr.AddErrorMsgs("For Z99ZZ99 ATC code, do not specify a combination code.");
                return;
            }

            combs = FetchAndConversion.listCombDdds.Where(comb => comb.ARS == pr.ARS).ToList();

            if (combs.Count > 0)
            {
                if (pr.CombinationObject == null)
                {
                    pr.SetValidate(false, "COMBINATION");
                    pr.AddErrorMsgs("This ATC code requires a Combination code, provide the Combination code. If no combination code exists for this combination, use Z99ZZ99_99.");
                }
                else
                {
                    bool found = combs.Any(comb => comb.Code == pr.CombinationObject.Code);

                    if (!found)
                    {
                        pr.SetValidate(false, "COMBINATION");
                        pr.AddErrorMsgs("The provided Combination code is not valid for this ATC code.");
                    }
                }
            }
            else
            {
                if (pr.CombinationObject != null)
                {
                    pr.SetValidate(false, "COMBINATION");
                    pr.AddErrorMsgs("This ATC code does not require a Combination code, remove the Combination code.");
                }
            }
        }

        private void ValidatePaediatrics(Product pr)
        {
            if (pr.PaediatricProduct == null)
            {
                pr.SetValidate(false, "PAEDIATRIC");
            }
            else
            {
                pr.SetValidate(true, "PAEDIATRIC");
            }
        }

        private void ValidateSector(Product pr)
        {
            pr.SetValidate(true, "SECTOR");

            if (string.IsNullOrEmpty(pr.Sector))
            {
                pr.SetValidate(false, "SECTOR");
                return;
            }
        }
        private void ValidateProductId(Product pr)
        {
            if (string.IsNullOrEmpty(pr.ProductId))
            {
                pr.SetValidate(false, "PROD_ID");
            }
            else
            {
                pr.SetValidate(true, "PROD_ID");
            }
        }

        private void ValidateLabel(Product pr)
        {
            if (string.IsNullOrEmpty(pr.Label)) // need to confirm Name or value
            {
                pr.SetValidate(false, "LABEL");
            }
            else
            {
                pr.SetValidate(true, "LABEL");
            }
        }

        private void ValidatePackageSize(Product pr)
        {
            if (!double.TryParse(pr.PackSize?.ToString(), out _)) // changed PACKSIZE to packsize doubtfull
            {
                pr.SetValidate(false, "PACK_SIZE");
            }
            else
            {
                pr.SetValidate(true, "PACK_SIZE");
            }
        }

        private void ValidateStrength(Product pr)
        {
            if (!double.TryParse(pr.Strength?.ToString(), out _)) // doubtful
            {
                pr.SetValidate(false, "STRENGTH");
                return;
            }

            if (pr.StrengthUnitObject == null)
            {
                pr.SetValidate(false, "STRENGTH");
                return;
            }

            pr.SetValidate(pr.StrengthUnitObject.IsValid, "STRENGTH");
        }

        private void ValidateDdds(Product pr)
        {
            dynamic WHODDD; // Equivalent to Variant in VBA
            object natDdd = null; // Placeholder if NatDDD is needed later
            object calcDdd = null; // Placeholder if CalcDDD is needed later

            // Check if `ars` is null or empty
            if (!pr.GetValidate("ARS"))
            {
                return;
            }

            // Determine WhoDdd based on CombinationObject
            if (pr.CombinationObject != null)
            {
                WHODDD = FetchAndConversion.FetchWhoCombinedDdd(pr.CombinationObject.Code);
            }
            else
            {
                WHODDD = FetchAndConversion.FetchWhoDdd(pr.ARS);
            }

            // Handle scenarios when WhoDdd is null
            if (WHODDD == null)
            {
                pr.AddInfoMsg("No DDD exists for the product, no number of DDDs will be calculated for this product.");
                pr.SetValidate(false, "DDD");
            }
            else
            {
                pr.WHODDD = WHODDD.DDD;
                pr.WHODDDUnit = WHODDD.DDD_Unit;
                pr.WHODDDUnitObject = FetchAndConversion.FetchUnit(WHODDD.DDD_Unit);
                pr.WhoDddValid = true;
                pr.SetValidate(true, "DDD");
            }
        }

        private void ValidateConcentration(Product pr)
        {
            // Initialize ConcentrationInUse to false
            pr.ConcentrationInUse = false;

            // Check if ConcentrationVolume is not null
            if (pr.ConcentrationVolume != null)
            {
                // ConcentrationVolume is provided
                pr.ConcentrationInUse = true;

                // Check if Volume is null
                if (pr.Volume == null)
                {
                    // ConcentrationVolume is provided but not Volume
                    pr.SetValidate(false, "CONCENTRATION");
                    pr.AddErrorMsgs("ConcentrationVolume has been provided, but not the corresponding Volume of bottle. The content cannot be calculated.");
                }
                else
                {
                    pr.SetValidate(true, "CONCENTRATION");
                }
            }
            else
            {
                // ConcentrationVolume is not provided but check if Volume is not null
                if (pr.Volume != null)
                {
                    // Volume is provided but ConcentrationVolume is missing
                    pr.ConcentrationInUse = true;
                    pr.AddErrorMsgs("Volume of bottle has been provided, but not the corresponding ConcentrationVolume. The content cannot be calculated.");
                    pr.SetValidate(false, "CONCENTRATION");
                }
            }
        }

        private void ValidateYearsAuthorizationWithdrawal(Product pr)
        {
            // Check if YearAuthorization or YearWithdrawal is null
            if (pr.YearAuthorization == null || pr.YearWithdrawal == null)
            {
                pr.SetValidate(true, "AUTH_WITHDR");
                return;
            }

            //Check if YearWithdrawal is earlier than YearAuthorization
            if (pr.YearWithdrawal < pr.YearAuthorization) // this need to check
            {
                pr.AddInfoMsg("Year of withdrawal cannot be before year of authorization.");
                pr.SetValidate(false, "AUTH_WITHDR");
            }
        }


        // Calculate Functions

        private void CalculateArs(Product pr)
        {
            string ars;

            if (pr.Atc5Object == null || pr.AdministrationRouteObject == null || pr.SaltObject == null)
            {
                pr.ARS = string.Empty;
                return;
            }
            else
            {
                ars = $"{pr.Atc5Object.Code}_{pr.AdministrationRouteObject.Code}{pr.SaltObject.Code}";
            }

            pr.ARS = ars;
            pr.SetValidate(true, "ARS");
        }

        public void CalculateAware(Product pr)
        {
            if (!pr.GetValidate("ARS"))
            {
                pr.SetValidate(false, "AwaRe");
                return;
            }
            if (pr.Atc5Object.Code == "Z99ZZ99" || pr.AtcAmClass != "ATB") // We set Not Applicable (N/A) for product with Z99ZZ99 code or not belonging to the ATB AM class
            {
                pr.AWaRe = new Aware { AWR = "N/A" };
                pr.SetValidate(true, "AwaRe");
                return;
            }
            // At this stage we only have valid ATB products
            // need refactor and having dynamic method to test same
            var awareList = listAware.Where(a => a.ATC5 == pr.ATC5).ToList();

            if (!awareList.Any()) // No AWR for ATC5, classify it as Not classified/undefined (N)
            {
                pr.AWaRe = new Aware { AWR = "N" };
                pr.SetValidate(true, "AwaRe");
                return;
            }
            if (awareList.Count == 1) // One AWR category for this ATC5
            {
                var aware = awareList.First();

                if (string.IsNullOrEmpty(aware.ROA)) // if the AWR category is defined for all routes 
                {
                    pr.AWaRe = aware;
                }
                else
                {
                    if (aware.ROA == pr.AdministrationRouteObject.Code)
                    {
                        pr.AWaRe = aware;
                    }
                    else // same ATC5 but not the same route, classify as Not classified/undefined
                    {
                        pr.AWaRe = new Aware { AWR = "N" };
                    }
                }
                pr.SetValidate(true, "AwaRe");
                return;
            }

            Aware matchedAware = null;
            foreach (var aware in awareList) // We have multiple AWR categories for the same ATC5
            {
                if (aware.ROA == pr.AdministrationRouteObject.Code)
                {
                    matchedAware = aware;
                    break;
                }
            }

            // If no match found, use fallback logic and assign the not classified/undefined AWR category
            if (matchedAware == null)
            {
                pr.AWaRe = new Aware { AWR = "N" };
            }
            else
            {
                pr.AWaRe = matchedAware;
            }
            pr.SetValidate(true, "AwaRe");
        }

        public void CalculateMEML(Product pr)
        {
            if (!pr.GetValidate("ARS"))
            {
                pr.SetValidate(false, "mEML");
                return;
            }
            if (pr.Atc5Object.Code == "Z99ZZ99")
            {
                pr.MEML = new MEML { mEML = "NO" };
                pr.SetValidate(true, "mEML");
                return;
            }

            // Get all matching MEML codes based on ATC5
            var matchingMEMLList = listmEML.Where(item => item.ATC5 == pr.Atc5Object.Code).ToList();

            if (!matchingMEMLList.Any()) // No EML category found for the ATC5
            {
                pr.MEML = new MEML { mEML = "NO" };
                pr.SetValidate(true, "mEML");
                return;
            }

            if (matchingMEMLList.Count == 1)
            {
                var mEML = matchingMEMLList.First();
                if (string.IsNullOrEmpty(mEML.ROA)) // This applies to all ROA
                {
                    pr.MEML = mEML;
                }
                else
                {
                    if (mEML.ROA == pr.AdministrationRouteObject.Code)
                    {
                        pr.MEML = mEML;
                    }
                    else
                    {
                        pr.MEML = new MEML { mEML = "NO" };
                    }
                }
                pr.SetValidate(true, "mEML");
                return;
            }

            // Multiple MEML codes, try to match ROA
            MEML matchedMEML = null;
            foreach (var mEML in matchingMEMLList) // We have multiple AWR categories for the same ATC5
            {
                if (mEML.ROA == pr.AdministrationRouteObject.Code)
                {
                    matchedMEML = mEML;
                    break;
                }
            }

            // If no match found, use fallback logic and assign the not classified/undefined AWR category
            if (matchedMEML == null)
            {
                pr.MEML = new MEML { mEML = "NO" };
            }
            else
            {
                pr.MEML = matchedMEML;
            }

            pr.SetValidate(true, "mEML");
        }

        private void CalculateConvFactor(Product pr)
        {

            if (!pr.GetValidate("ARS"))
            {
                pr.ConversionFactor = null;
                return;
            }

            if (pr.ATC5 == "Z99ZZ99" || (pr.CombinationObject != null && pr.CombinationObject.Code == "Z99ZZ99_99"))
            {
                pr.ConversionFactor = 1;
                pr.SetValidate(true, "CONV_FACTOR");
                return;
            }

            if (!pr.GetValidate("STRENGTH") || !pr.GetValidate("DDD"))
            {
                pr.ConversionFactor = null;
                pr.SetValidate(false, "CONV_FACTOR");
                return;
            }

            MeasureUnit su = FetchAndConversion.FetchUnit(pr.StrengthUnitObject.Value.Code);
            MeasureUnit du = FetchAndConversion.FetchUnit(pr.WHODDDUnit);

            if (su.Family == du.Family)
            {
                pr.ConversionFactor = 1;
                pr.SetValidate(true, "CONV_FACTOR");
                return;
            }
            
            ConversionFactor matchedConvFactor = null;
            foreach (var cf in listConvFactors)
            {
                if (cf.ATC5 != pr.Atc5Object.Code || cf.AdministrationRoute != pr.AdministrationRouteObject.Code)
                {
                    continue;
                }
                var cf_fromUnitFam = FetchAndConversion.FetchUnit(cf.Unit_From)?.Family;
                var cf_ToUnitFam = FetchAndConversion.FetchUnit(cf.Unit_To)?.Family;
                if (su.Family != cf_fromUnitFam || du.Family != cf_ToUnitFam)
                {
                    continue;
                }
                matchedConvFactor = cf;
                break;
            }

            if (matchedConvFactor == null)
            {
                pr.ConversionFactor = null;
                pr.SetValidate(false, "CONV_FACTOR");
                return;
            }

            pr.ConversionFactor = matchedConvFactor.Factor;
            if (pr.ConversionFactor == null)
            {
                pr.SetValidate(false, "CONV_FACTOR");
            }
            else
            {
                pr.SetValidate(true, "CONV_FACTOR");
            }
        }

        private void CalculatePackageContent(Product pr)
        {
            double? stdStrength;
            double? stdConcentrationVolume;
            double? stdVolume;
            MeasureUnit strgUnit;
            MeasureUnit strgUnit2;
            double PackContent;

            // Check if ars is null or empty
            if (!pr.GetValidate("ARS"))
            {
                return;
            }

            // Check for specific conditions that require exiting
            if (pr.ATC5 == "Z99ZZ99" || (pr.CombinationObject != null && pr.CombinationObject.Code == "Z99ZZ99_99")) // need to check Ats5
            {
                return;
            }

            // Validate various properties
            if (!pr.GetValidate("PACK_SIZE"))
            {
                pr.AddInfoMsg("Package content cannot be calculated because package size is not valid.");
                return;
            }

            if (!pr.GetValidate("STRENGTH"))
            {
                pr.AddInfoMsg("Package content cannot be calculated because strength is not valid.");
                return;
            }

            if (pr.ConcentrationInUse && !pr.GetValidate("CONCENTRATION"))
            {
                pr.AddInfoMsg("Package content cannot be calculated because either concentration volume or volume is invalid.");
                return;
            }


            stdStrength = ConvertBaseUnit(pr.Strength.Value, pr.StrengthUnitObject.Value.Code); // might be break

            if (!pr.ConcentrationInUse)
            {
                stdConcentrationVolume = 1;
                stdVolume = 1;
            }
            else
            {
                stdConcentrationVolume = pr.ConcentrationVolume.Value;
                stdVolume = pr.Volume.Value;
            }

            PackContent = (double)((pr.PackSize * stdStrength * stdVolume / stdConcentrationVolume)); // chanage pr.PackSize to pr.PackSize.value

            pr.PackContent = PackContent;

            // Determine the unit based on strength unit family

            strgUnit = FetchAndConversion.FetchUnit(pr.StrengthUnitObject.Value.Code);
            if (strgUnit.Family == "GRAM")
            {
                strgUnit2 = FetchAndConversion.FetchUnit("G");
                pr.PackContentUnit = "G";
                pr.PackContentUnitObject = strgUnit2;
            }
            else if (strgUnit.Family == "UNIT_DOSE")
            {
                strgUnit2 = FetchAndConversion.FetchUnit("UD");
                pr.PackContentUnit = "UD";
                pr.PackContentUnitObject = strgUnit2;
            }
            else
            {
                strgUnit2 = FetchAndConversion.FetchUnit("MU");
                pr.PackContentUnit = "MU";
                pr.PackContentUnitObject = strgUnit2;
            }
            pr.SetValidate(true, "PACK_CONTENT");
        }

        private void CalculateDddPerPackage(Product pr)
        {
            double stdDdd;
            double dpp;

            // Check if ARS is null, empty, or if DDD is not validated, or if conversion factor is not validated
            if (!pr.GetValidate("ARS") || !pr.GetValidate("DDD") || !pr.GetValidate("CONV_FACTOR"))
            {
                return;
            }

            // Exit if ATC5 or Combination match the specified conditions
            if (pr.ATC5 == "Z99ZZ99" || (pr.CombinationObject != null && pr.CombinationObject.Code == "Z99ZZ99_99"))
            {
                return;
            }

            // Convert base unit for DDD
            stdDdd = (double)ConvertBaseUnit(pr.WHODDD, pr.WHODDDUnit);


            // Calculate DDD per package
            if (stdDdd == 0)
            {
                dpp = 0.0;
            }
            else
            {
                dpp = (double)(pr.PackContent * pr.ConversionFactor / stdDdd);
            }

            // Assign calculated value to DDD per package
            pr.DPP = dpp;

            // Set validation flag for DPP
            pr.SetValidate(true, "DPP");
        }

        // Utility functions

        public static double? ConvertBaseUnit(double amount, string unitMeasure)
        {
            // Fetch the unit from the list
            MeasureUnit u = listUnits.FirstOrDefault(item => item.Code == unitMeasure);

            if (u != null)
            {
                double? newAmount = amount * u.BaseConversion;
                return newAmount;
            }
            return null;
        }
*/
    }
}