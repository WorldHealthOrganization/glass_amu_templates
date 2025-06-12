// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;
using AMU_Template.Constants;
using AMU_Template.Helpers;
using AMU_Template.Models;
using NAMU_Template.Constants;
using NAMU_Template.Data_Validation;
using NAMU_Template.Helper;

namespace NAMU_Template.Models
{
    public class Product : Medicine
    {

        // WHO Excel Template Fields

        public const string UID_FIELD = "UID";
        public const string COUNTRY_FIELD = "COUNTRY";
        public const string PRODUCT_ID_FIELD = "PRODUCT_ID";
        public const string LABEL_FIELD = "LABEL";
        public const string PACKSIZE_FIELD = "PACKSIZE";
        public const string ROUTE_ADMIN_FIELD = "ROUTE_ADMIN";
        public const string STRENGTH_FIELD = "STRENGTH";
        public const string STRENGTH_UNIT_FIELD = "STRENGTH_UNIT";
        public const string CONCENTRATION_VOLUME_FIELD = "CONCENTRATION_VOLUME";
        public const string VOLUME_FIELD = "VOLUME";
        public const string ATC5_FIELD = "ATC5";
        public const string SALT_FIELD = "SALT";
        public const string COMBINATION_FIELD = "COMBINATION";
        public const string PAEDIATRIC_FIELD = "PAEDIATRIC_PRODUCT";
        public const string FORM_FIELD = "FORM";
        public const string PRODUCT_NAME_FIELD = "PRODUCT_NAME";
        public const string INGREDIENTS_FIELD = "INGREDIENTS";
        public const string PRODUCT_ORIGIN_FIELD = "PRODUCT_ORIGIN";
        public const string MANUFACTURER_COUNTRY_FIELD = "MANUFACTURER_COUNTRY";
        public const string MARKET_AUTH_HOLDER_FIELD = "MARKET_AUTH_HOLDER";
        public const string GENERICS_FIELD = "GENERICS";
        public const string YEAR_AUTHORIZATION_FIELD = "YEAR_AUTHORIZATION";
        public const string YEAR_WITHDRAWAL_FIELD = "YEAR_WITHDRAWAL";
        public const string SECTOR_FIELD = "SECTOR";

        // WIDP Excel Template Fields
        public const string CONCENTRATION_VOLUME_UNIT_WIDP_FIELD = "CONCENTRATION_VOLUME_UNIT_WIDP_FIELD";
        public const string VOLUME_UNIT_WIDP_FIELD = "VOLUME_UNIT_WIDP_FIELD";
        public const string DATA_STATUS_WIDP_FIELD = "DATA_STATUS_WIDP_FIELD";
        public const string ENROLMENT_DATE_WIDP_FIELD = "ENROLMENT_DATE_WIDP_FIELD";
        public const string INCIDENT_DATE_WIDP_FIELD = "INCIDENT_DATE_WIDP_FIELD";


        // Validation constants

        public const string COUNTRY_VALIDATION = "COUNTRY_VALIDATION";
        public const string PRODUCT_ID_VALIDATION = "PRODUCT_ID_VALIDATION";
        public const string LABEL_VALIDATION = "LABEL_VALIDATION";
        public const string PACKSIZE_VALIDATION = "PACKSIZE_VALIDATION";
        public const string STRENGTH_VALIDATION = "STRENGTH_VALIDATION";
        public const string ROA_VALIDATION = "ROA_VALIDATION";
        public const string ATC5_VALIDATION = "ATC5_VALIDATION";
        public const string SALT_VALIDATION = "SALT_VALIDATION";
        public const string COMBINATION_VALIDATION = "COMBINATION_VALIDATION";
        public const string CONVERSION_VALIDATION = "CONVERSION_VALIDATION";
        public const string CONCENTRATION_VALIDATION = "CONCENTRATION_VALIDATION";
        public const string DDD_VALIDATION = "DDD_VALIDATION";
        public const string PAEDIATRIC_VALIDATION = "PAEDIATRIC_VALIDATION";
        public const string CONTENT_VALIDATION = "CONTENT_VALIDATION";
        public const string ARS_VALIDATION = "ARS_VALIDATION";
        public const string MANUFACTURER_COUNTRY_VALIDATION = "MANUFACTURER_COUNTRY_VALIDATION";
        public const string AUTH_WITH_VALIDATION = "AUTH_WITH_VALIDATION";
        public const string SECTOR_VALIDATION = "SECTOR_VALIDATION";
        public const string DPP_VALIDATION = "DPP_VALIDATION";
        public const string AWR_VALIDATION = "AWR_VALIDATION";
        public const string MEML_VALIDATION = "MEML_VALIDATION";
        public const string ATCCLASS_VALIDATION = "ATCCLASS_VALIDATION";


        public Dictionary<string, IDataField?> DataFieldValues;

        public int ProductLineNo { get; set; }

        public string Country { get; set; }

        public int Year { get; set; }

        public string ProductId { get; set; }

        public Decimal PackSize { get; set; }

        public Decimal Strength { get; set; }

        public MeasureUnit StrengthUnit { get; set; }

        public Decimal? ConcentrationVolume { get; set; }

        public Decimal? Volume { get; set; }

        public ATC ATC4 { get; set; }

        public ATC ATC3 { get; set; }

        public ATC ATC2 { get; set; }

        public Salt Salt { get; set; }

        public DDDCombination Combination { get; set; }

        public YesNoUnknown Paediatric { get; set; }

        public string AWaRe { get; set; }

        public YesNoNA MEML { get; set; }

        public HealthSector Sector { get; set; }    

        public string Form { get; set; }

        public string ProductName { get; set; }

        public string Ingredients { get; set; }

        public string ProductOrigin { get; set; }

        public string ManufacturerCountry { get; set; }

        public string MAHolder { get; set; }

        public YesNoUnknown Generics { get; set; }

        public int YearAuthorization { get; set; }

        public int YearWithdrawal { get; set; }

        public Decimal ConversionFactor;

        private string InternalCachedProductUniqueId;

        public override string UniqueId
        {
            get
            { if (InternalCachedProductUniqueId != null)
                {
                    return InternalCachedProductUniqueId;
                }
                else
                {
                    if (ProductLineNo == 0 || string.IsNullOrEmpty(ProductId))
                    {
                        return "UNDEFINED";
                    }
                    else
                    {
                        InternalCachedProductUniqueId = $"{ProductLineNo}|{ProductId}";
                        return InternalCachedProductUniqueId;
                    }
                }
            }
        }

        public int SequenceNo { get; set; }

        public List<string> Errors { get; set; } = new List<string>();

        public List<string> Warnings { get; set; } = new List<string>();

        public List<string> Infos { get; set; } = new List<string>();

        public EntityStatus Status { get; set; }

        public Dictionary<string, bool> Validations;

        public List<ValidationMessage> ValidationMessages { get; set; }


        public Product()
        {
            Errors = new List<string>();
            Warnings = new List<string>();
            Infos = new List<string>();
            Status = EntityStatus.OK; // Assuming STATUS_OK is a string constant

            DataFieldValues = new Dictionary<string, IDataField?>
            { 
                {COUNTRY_FIELD, null}, 
                {PRODUCT_ID_FIELD, null}, 
                {LABEL_FIELD, null}, 
                {PACKSIZE_FIELD, null}, 
                {ROUTE_ADMIN_FIELD, null},
                {STRENGTH_FIELD, null}, 
                {STRENGTH_UNIT_FIELD, null}, 
                {CONCENTRATION_VOLUME_FIELD, null}, 
                {VOLUME_FIELD, null}, 
                {ATC5_FIELD, null},
                {SALT_FIELD, null}, 
                {COMBINATION_FIELD, null}, 
                {PAEDIATRIC_FIELD, null}, 
                {FORM_FIELD, null}, 
                {PRODUCT_NAME_FIELD, null},
                {INGREDIENTS_FIELD, null}, 
                {PRODUCT_ORIGIN_FIELD, null}, 
                {MANUFACTURER_COUNTRY_FIELD, null}, 
                {MARKET_AUTH_HOLDER_FIELD, null},
                {GENERICS_FIELD, null}, 
                {YEAR_AUTHORIZATION_FIELD, null}, 
                {YEAR_WITHDRAWAL_FIELD, null }, 
                {SECTOR_FIELD, null},
            };

            ValidationMessages = new List<ValidationMessage>(); // Initialize the collection before adding items
            //Validations = new Dictionary<string, bool>();
            // Needs to check this on load data
            Validations = new Dictionary<string, bool>
                {
                { COUNTRY_VALIDATION, false },
                { PRODUCT_ID_VALIDATION, false },
                { LABEL_VALIDATION, false },
                { PACKSIZE_VALIDATION, false },
                { STRENGTH_VALIDATION, false },
                { ROA_VALIDATION, false },
                { ATC5_VALIDATION, false },
                { SALT_VALIDATION, false }, 
                { COMBINATION_VALIDATION, false },
                { CONCENTRATION_VALIDATION, false },
                { CONVERSION_VALIDATION, false },
                { DDD_VALIDATION, false },
                { PAEDIATRIC_VALIDATION, false },
                { CONTENT_VALIDATION, false },
                { ARS_VALIDATION, false },
                { MANUFACTURER_COUNTRY_VALIDATION, false },
                { AUTH_WITH_VALIDATION, false },
                { SECTOR_VALIDATION, false },
                { DPP_VALIDATION, false },
                { AWR_VALIDATION, false },
                { MEML_VALIDATION, false },
                { ATCCLASS_VALIDATION, true },
            };
        }

        ~Product()
        {
            Errors = null;
            Infos = null;
            Status = EntityStatus.DEFAULT;
        }

        public void SetField(string fieldName, IDataField dataField)
        {
            DataFieldValues[fieldName] = dataField;
        }

        public void UpdateStatus(int level)
        {
            if ((int)Status < level)
            {
                Status = (EntityStatus)level;
            }
        }
        public string GetStatusMessages()
        {
            List<string> msgs = new List<string>();

            for (int i = 0; i < ValidationMessages.Count; i++)
            {
                if (ValidationMessages[i].Message != null)
                    msgs.Add($"{ValidationMessages[i].MessageType}: {ValidationMessages[i].Message}");
            }
            return string.Join(Environment.NewLine, msgs);
        }
        public void AddErrorMsgss(string msg, object dataField)
        {
            ValidationMessages.Add(new ValidationMessage() { MessageType = ValidationMessageType.Error, Message = msg, ErrorField = dataField });
            this.Status = EntityStatus.ERROR;
        }

        public void AddErrorMsgs(string data)
        {
            ValidationMessages.Add(new ValidationMessage() { MessageType = ValidationMessageType.Error, Message = data });
            this.Status = EntityStatus.ERROR;
        }

        public void AddInfoMsg(string msg, dynamic dataField)
        {
            ValidationMessages.Add(new ValidationMessage() { MessageType = ValidationMessageType.Info, Message = msg, ErrorField = dataField });
            //If the status is already Error then don't update it..!
            if (this.Status != EntityStatus.ERROR)
                this.Status = EntityStatus.INFO;
        }

        public void AddInfoMsg(dynamic data)
        {
            if (data is string)
            {
                ValidationMessages.Add(new ValidationMessage() { MessageType = ValidationMessageType.Info, Message = data });
            }
            else
            {
                ValidationMessages.Add(new ValidationMessage() { MessageType = ValidationMessageType.Info, ErrorField = data });
            }

            //If the status is already Error then don't update it..!
            if (this.Status != EntityStatus.ERROR)
                this.Status = EntityStatus.INFO;
        }

        public void AddWarningMsg(string msg)
        {
            ValidationMessages.Add(new ValidationMessage()
            {
                MessageType = ValidationMessageType.Warning,
                Message = msg
            });

            // If the status is already Error, don't downgrade it to Warning.
            if (this.Status != EntityStatus.ERROR)
                this.Status = EntityStatus.WARNING;
        }

        // getValidate

        public bool IsProductStatusError()
        {
            return Errors.Count > 0;
        }

        public bool IsProductValid()
        {
            foreach (var validation in Validations)
            {
                if (!validation.Value && validation.Key!=DDD_VALIDATION) // DDD validation should not be counted in the overall validation.
                {
                    return false;
                }
            }
            return true;
        }

        #region Validation

        public void ValidateProduct(bool force)
        {
            if (this.Infos.Count > 0)
            {
                this.Status = EntityStatus.INFO;
            }

            if (this.Warnings.Count > 0)
            {
                this.Status = EntityStatus.WARNING;
            }

            if (this.Errors.Count > 0)
            {
                this.Status = EntityStatus.ERROR;
            }

            if (this.Status > EntityStatus.OK && !force)
            {
                return;
            }

            ValidateOptionalFields();
            ValidateCountry();
            ValidateProductId();
            ValidateLabel();
            ValidateAtc5();
            ValidateRoA();
            ValidateSalt();
            ValidateArs();
            ValidateCombination();
            ValidatePackSize();
            ValidateStrength();
            ValidateConcentration();
            ValidatePaediatrics();
            ValidateManufacturerCountry();
            ValidateYearsAuthorizationWithdrawal();
            ValidateDdds();
            ValidateSector();
            ValidateConvFactor();
            ValidatePackageContent();
            ValidateDddPerPackage();
            ValidateAware();
            ValidateMEML();
            FinalizeValidation();
        }

        private void ValidateOptionalFields()
        {
            DataField<string> dfForm = (DataField<string>)DataFieldValues[Product.FORM_FIELD];
            if (dfForm.IsValid && !dfForm.IsMissing)
            {
                this.Form = dfForm.Value;
            }
            DataField<string> dfName = (DataField<string>)DataFieldValues[Product.PRODUCT_NAME_FIELD];
            if (dfName.IsValid && !dfName.IsMissing)
            {
                this.ProductName = dfName.Value;
            }
            DataField<string> dfIngrs = (DataField<string>)DataFieldValues[Product.INGREDIENTS_FIELD];
            if (dfIngrs.IsValid && !dfIngrs.IsMissing)
            {
                this.Ingredients = dfIngrs.Value;
            }
            DataField<string> dfOrig = (DataField<string>)DataFieldValues[Product.PRODUCT_ORIGIN_FIELD];
            if (dfOrig.IsValid && !dfOrig.IsMissing)
            {
                this.ProductOrigin = dfOrig.Value;
            }
            //DataField<string> dfCoOrig = (DataField<string>)DataFieldValues[Product.MANUFACTURER_COUNTRY_FIELD];
            //if (dfCoOrig.IsValid && !dfCoOrig.IsMissing)
            //{
            //    this.ManufacturerCountry = dfCoOrig.Value;
            //}
            DataField<string> dfMah = (DataField<string>)DataFieldValues[Product.MARKET_AUTH_HOLDER_FIELD];
            if (dfMah.IsValid && !dfMah.IsMissing)
            {
                this.MAHolder = dfMah.Value;
            }
            DataField<YesNoUnknown> dfGen = (DataField<YesNoUnknown>)DataFieldValues[Product.GENERICS_FIELD];
            if (dfGen.IsValid && !dfGen.IsMissing)
            {
                this.Generics = dfGen.Value;
            }
            DataField<int> dfYA = (DataField<int>)DataFieldValues[Product.YEAR_AUTHORIZATION_FIELD];
            if (dfYA.IsValid && !dfYA.IsMissing)
            {
                this.YearAuthorization = dfYA.Value;
            }
            DataField<int> dfYW = (DataField<int>)DataFieldValues[Product.YEAR_WITHDRAWAL_FIELD];
            if (dfYW.IsValid && !dfYW.IsMissing)
            {
                this.YearWithdrawal = dfYW.Value;
            }
        }

        private void FinalizeValidation()
        {
           
        }

        private void ValidateCountry()
        {
            DataField<string> cntry = (DataField<string>)DataFieldValues[COUNTRY_FIELD];

            if (!cntry.IsValid)
            {
                SetValidate(false, COUNTRY_VALIDATION);
                return;
            }
            if (ThisWorkbook.CountryDataDict.ContainsKey(cntry.Value))
            {
                this.Country = cntry.Value;
                SetValidate(true, COUNTRY_VALIDATION);
            }
            else
            {
                {
                    SetValidate(false, COUNTRY_VALIDATION);
                    AddErrorMsgs("Country is not in the valid list of country codes.");
                }
            }
        }

        private void ValidateProductId()
        {
            if (IsValidDataField(PRODUCT_ID_FIELD))
            {
                this.ProductId = (string)DataFieldValues[PRODUCT_ID_FIELD].Value;
                SetValidate(true, PRODUCT_ID_VALIDATION);
            }
            else { SetValidate(false, PRODUCT_ID_VALIDATION); }
        }

        private void ValidateLabel()
        {
            if (IsValidDataField(LABEL_FIELD))
            {
                this.Label = (string)DataFieldValues[LABEL_FIELD].Value;
                SetValidate(true, LABEL_VALIDATION);
            }
            else { SetValidate(false, LABEL_VALIDATION); }
        }

        private void ValidateAtc5()
        {
            var atc5_data = DataFieldValues[ATC5_FIELD];
            if (!atc5_data.IsValid)
            {
                SetValidate(false, ATC5_VALIDATION);
                if (atc5_data.IsMissing)
                {
                    AddErrorMsgs($"If there is no ATC5 for this product, use the code {AMUConstants.ATC_Z99_CODE}.");
                }
                else
                {
                    AddErrorMsgs($"ATC5 code {atc5_data.InputValue} is not valid.");
                }
                return;
            }
            
            this.ATC5 = (ATC)atc5_data.Value;
            this.ATC4 = Utils.GetATCParent(this.ATC5);
            this.ATC3 = Utils.GetATCParent(this.ATC4);
            this.ATC2 = Utils.GetATCParent(this.ATC3);
            this.AMClass = this.ATC5.AMClass;
            this.ATCClass = this.ATC5.ATCClass;
            SetValidate(true, ATC5_VALIDATION);
        }

        private void ValidateRoA()
        {
            if (IsValidDataField(ROUTE_ADMIN_FIELD))
            {
                this.Roa = (AdministrationRoute)DataFieldValues[ROUTE_ADMIN_FIELD].Value;
                SetValidate(true, ROA_VALIDATION);
            }
            else { SetValidate(false, ROA_VALIDATION); }
        }

        private void ValidateSalt()
        {
            if (IsValidDataField(SALT_FIELD))
            {
                Salt salt = (Salt)DataFieldValues[SALT_FIELD].Value;
                if (salt.Code != "XXXX")
                {
                    if (GetValidate(ATC5_VALIDATION))
                    {
                        if (!salt.Atc5s.Contains(this.ATC5.Code))
                        {
                            AddErrorMsgs($"The salt {salt.Code} is not applicable to ATC5 {this.ATC5.Code}.");
                            SetValidate(false, SALT_VALIDATION);
                            return;
                        }
                        this.Salt = (Salt)DataFieldValues[SALT_FIELD].Value;
                        SetValidate(true, SALT_VALIDATION);
                        return;
                    }
                }
                this.Salt = (Salt)DataFieldValues[SALT_FIELD].Value;
                SetValidate(true, SALT_VALIDATION);
            }
            else { SetValidate(false, SALT_VALIDATION); }
        }

        private void ValidateArs()
        {
            if (!IsMandatoryValidationValid(ATC5_VALIDATION) ||
                !IsMandatoryValidationValid(ROA_VALIDATION) ||
                !IsMandatoryValidationValid(SALT_VALIDATION))
            {
                SetValidate(false, ARS_VALIDATION);
                return;
            }
            this.ARS = ARSHelper.GenerateARSFromATC5ROASalt(ATC5.Code, Roa.Code, Salt.Code);
            SetValidate(true, ARS_VALIDATION);
        }

        private void ValidateCombination()
        {
            if (!IsMandatoryValidationValid(ATC5_VALIDATION))
            {
                return;
            }
            if (IsValidDataField(COMBINATION_FIELD))
            {
                if (DataFieldValues[COMBINATION_FIELD].IsMissing)
                {
                    SetValidate(true, COMBINATION_VALIDATION);
                    return;
                }
                DDDCombination comb = (DDDCombination)DataFieldValues[COMBINATION_FIELD].Value;
                
                if (this.ATC5.Code == AMUConstants.ATC_Z99_CODE && comb.Code != AMUConstants.COMB_Z99_CODE)
                {
                    AddErrorMsgs($"The ATC5 code is defined as {AMUConstants.ATC_Z99_CODE} code, the combination code must be {AMUConstants.COMB_Z99_CODE}.");
                    SetValidate(false, COMBINATION_VALIDATION);
                    return;
                }
                if (comb.Code == AMUConstants.COMB_Z99_CODE && String.IsNullOrEmpty(Ingredients))
                {
                    AddErrorMsgs($"The Combination code is defined as {AMUConstants.COMB_Z99_CODE}. Please provide the list of ingredients as INN separated by comma in the ingredients column.");
                    AddInfoMsg("Ensure that the ingredients and their respective strength are stated in the label.");
                    SetValidate(false, COMBINATION_VALIDATION);
                    return;
                }
                if (this.ATC5.Code != AMUConstants.ATC_Z99_CODE && comb.Code !=AMUConstants.COMB_Z99_CODE && comb.ATC5.Code != this.ATC5.Code)
                {
                    AddErrorMsgs($"The Combination code {comb.Code} and ATC code {this.ATC5.Code} do not match.");
                    SetValidate(false, COMBINATION_VALIDATION);
                    return;
                }
                this.Combination = comb;
                SetValidate(true, COMBINATION_VALIDATION);
                if (this.ATC5.Code != AMUConstants.ATC_Z99_CODE && comb.Code == AMUConstants.COMB_Z99_CODE)
                {
                    AddInfoMsg($"The undefined combination code is within the ATC code {this.ATC5.Code}.");
                }
            }
            else
            {
                SetValidate(false, COMBINATION_VALIDATION);
            }
        }

        private void ValidatePackSize()
        {
            if (IsValidDataField(PACKSIZE_FIELD))
            {
                var ps = (Decimal)DataFieldValues[PACKSIZE_FIELD].Value;
                if (ps <=0) 
                {
                    SetValidate(false, PACKSIZE_VALIDATION);
                    AddErrorMsgs("PACKSIZE must be positive.");
                }
                else
                {
                    this.PackSize = ps;
                    SetValidate(true, PACKSIZE_VALIDATION);
                }

            }
            else { SetValidate(false, PACKSIZE_VALIDATION); }
        }

        private void ValidateStrength()
        {
            DataField<Decimal> str_data = (DataField<Decimal>)DataFieldValues[STRENGTH_FIELD];
            DataField<MeasureUnit> uni_data = (DataField<MeasureUnit>)DataFieldValues[STRENGTH_UNIT_FIELD];
            if (str_data.IsValid && uni_data.IsValid)
            {
                if (str_data.Value < 0)
                {
                    SetValidate(false, STRENGTH_VALIDATION);
                    AddErrorMsgs("STRENGTH must be positive");
                    return;
                }
                if (str_data.Value == 0)
                {
                    SetValidate(true, STRENGTH_VALIDATION);
                    AddInfoMsg("STRENGTH is 0. Calculated amount will be 0.");
                    return;
                }

                this.Strength = str_data.Value;
                this.StrengthUnit = uni_data.Value;
                SetValidate(true, STRENGTH_VALIDATION);
            }
            else
            {
                SetValidate(false, STRENGTH_VALIDATION);
                
            }
        }

        private void ValidateConcentration()
        {
            if(!IsMandatoryValidationValid(STRENGTH_VALIDATION))
            {
                return;
            }
            DataField<Decimal> conc_data = (DataField<Decimal>)DataFieldValues[CONCENTRATION_VOLUME_FIELD];
            DataField<Decimal> vol_data = (DataField<Decimal>)DataFieldValues[VOLUME_FIELD];
            if (conc_data.IsValid  && vol_data.IsValid)
            {
                if (conc_data.IsMissing && vol_data.IsMissing)
                {
                    SetValidate(true, CONCENTRATION_VALIDATION);
                    return;
                }
                if (conc_data.IsProvided && vol_data.IsMissing)
                {
                    SetValidate(false, CONCENTRATION_VALIDATION);
                    AddErrorMsgs("CONCENTRATION_VOLUME is provided but VOLUME is not.");
                    return;
                }
                if (conc_data.IsMissing && vol_data.IsProvided)
                {
                    SetValidate(false, CONCENTRATION_VALIDATION);
                    AddErrorMsgs("VOLUME is provided but CONCENTRATION_VOLUME is not.");
                    return;
                }
                if (conc_data.Value <= 0)
                {
                    SetValidate(false, CONCENTRATION_VALIDATION);
                    AddErrorMsgs("CONCENTRATION VOLUME must be positive");
                    return;
                }
                ConcentrationVolume = conc_data.Value;
                Volume = vol_data.Value;
                SetValidate(true, CONCENTRATION_VALIDATION); 
            }
            else
            {
                SetValidate(false, CONCENTRATION_VALIDATION);
                return;
            }
        }

        private void ValidateManufacturerCountry()
        {
            DataField<string> manCntry = (DataField<string>)DataFieldValues[MANUFACTURER_COUNTRY_FIELD];
            
            if (!manCntry.IsValid)
            {
                SetValidate(false, MANUFACTURER_COUNTRY_VALIDATION);
                return;
            }
            if (manCntry.IsMissing)
            {
                SetValidate(true, MANUFACTURER_COUNTRY_VALIDATION);
                return;
            }
            if (ThisWorkbook.CountryDataDict.ContainsKey(manCntry.Value))
            {
                if (IsMandatoryValidationValid(COUNTRY_VALIDATION) && DataFieldValues[PRODUCT_ORIGIN_FIELD].IsValid && !DataFieldValues[PRODUCT_ORIGIN_FIELD].IsMissing && (string)DataFieldValues[PRODUCT_ORIGIN_FIELD].Value=="DOMPROD") 
                {
                    if (this.Country != (string)DataFieldValues[MANUFACTURER_COUNTRY_FIELD].Value)
                    {
                        SetValidate(false, MANUFACTURER_COUNTRY_VALIDATION);
                        AddErrorMsgs("Manufacturer Country should be the same as country for domestic production.");
                        return;
                    }
                }
                this.ManufacturerCountry = manCntry.Value;
                SetValidate(true, MANUFACTURER_COUNTRY_VALIDATION);
            }
            else
            {
                {
                    SetValidate(false, MANUFACTURER_COUNTRY_VALIDATION);
                    AddErrorMsgs("Manufacturer Country is not in the valid list of country codes.");
                }
            }
        }

        private void ValidateYearsAuthorizationWithdrawal()
        {
            DataField<int> yAuth = (DataField<int>)DataFieldValues[YEAR_AUTHORIZATION_FIELD];
            DataField<int> yWith = (DataField<int>)DataFieldValues[YEAR_WITHDRAWAL_FIELD];
            if (!yAuth.IsValid || !yWith.IsValid)
            {
                SetValidate(false, AUTH_WITH_VALIDATION);
                return;
            }
            if ((yAuth.IsValid && yWith.IsValid) && (yAuth.IsMissing && yWith.IsMissing))
            {
                SetValidate(true, AUTH_WITH_VALIDATION);
                return;
            }
                if (!yAuth.IsMissing && !yWith.IsMissing)
            {
                bool res = yAuth.Value <= yWith.Value;
                if (res == false)
                {
                    SetValidate(false, AUTH_WITH_VALIDATION);
                    AddErrorMsgs("Year of withdrawal must be greater or equal than Year of authorization.");
                }
                else
                {
                    SetValidate(true, AUTH_WITH_VALIDATION);
                }
            }
        }

        private void ValidateDdds()
        {
            if (!IsMandatoryValidationValid(ARS_VALIDATION) || ! IsMandatoryValidationValid(COMBINATION_VALIDATION))
            {
                return;
            }
            var dfComb = (DataField<DDDCombination>)DataFieldValues[COMBINATION_FIELD];
            if (dfComb.IsValid)
            {
                if (!dfComb.IsMissing)
                { // we have combination, then use it for DDD
                    Amount dddAmount = new Amount(this.Combination.DDDValue, this.Combination.DDDUnit);
                    this.DDD = dddAmount;
                    SetValidate(true, DDD_VALIDATION);
                    return;
                }
            }
            else
            {
                // we have an invalid combination
                SetValidate(false, DDD_VALIDATION);
                return;
            }
                
            // we don't have combination, then check we have a valid DDD for ARS
            if (ThisWorkbook.DDDDataDict.ContainsKey(ARS))
            {
                var ddd = ThisWorkbook.DDDDataDict[ARS];
                Amount dddAmount = new Amount(ddd.Value, ddd.Unit);
                this.DDD = dddAmount;
                SetValidate(true, DDD_VALIDATION);
                return;
            }
            else
            {
                // We don't have DDD for this ARS
                SetValidate(false, DDD_VALIDATION);
                AddInfoMsg("No DDD exists, no number of DDDs will be calculated for this product.");
                return;
            }
        }

        private void ValidateSector()
        {
            if (IsValidDataField(SECTOR_FIELD))
            {
                Sector = (HealthSector)DataFieldValues[SECTOR_FIELD].Value;
                SetValidate(true, SECTOR_VALIDATION);
            }
            else { SetValidate(false, SECTOR_VALIDATION); }
        }

        private void ValidatePaediatrics()
        {
            if (IsValidDataField(PAEDIATRIC_FIELD))
            {
                SetValidate(true, PAEDIATRIC_VALIDATION);
                Paediatric = (YesNoUnknown)DataFieldValues[PAEDIATRIC_FIELD].Value;
            }
            else { SetValidate(false, PAEDIATRIC_VALIDATION); }
        }

        private void ValidateConvFactor()
        {
            if (!IsMandatoryValidationValid(ARS_VALIDATION) || !IsMandatoryValidationValid(STRENGTH_VALIDATION) || !IsMandatoryValidationValid(DDD_VALIDATION))
            {
                return;
            }

            if (this.ATC5.Code == AMUConstants.ATC_Z99_CODE || (this.Combination?.Code == AMUConstants.COMB_Z99_CODE))
            {
                return;
            }

            if (this.DDD.Unit == this.StrengthUnit)
            {
                this.ConversionFactor = Decimal.One;
                SetValidate(true, CONVERSION_VALIDATION);
                return;
            }

            if (this.DDD.Unit.Family == this.StrengthUnit.Family)
            {
                this.ConversionFactor = Decimal.One;
                SetValidate(true, CONVERSION_VALIDATION);
                return;
            }

            var factorList = ThisWorkbook.ConversionFactorDataList;
            ConversionFactor? cf = factorList.FirstOrDefault(f => this.ARS == f.ARS && this.StrengthUnit.Family == f.UnitFrom.Family && this.DDD.Unit.Family == f.UnitTo.Family);
            if (cf != null)
            {
                this.ConversionFactor = cf.Factor;
                SetValidate(true, CONVERSION_VALIDATION );
            }
            else 
            {
                AddErrorMsgs($"Units of strength {this.StrengthUnit.Code} and DDD {this.DDD.Unit.Code} are incompatible");
                SetValidate(false, CONVERSION_VALIDATION);
                return;
            }
        }

        private void ValidatePackageContent()
        {
            bool prevalid = true;
            Amount sAmount;

            // Check if ars is null or empty
            if (!IsMandatoryValidationValid(ARS_VALIDATION))
            {
                AddInfoMsg("Package content cannot be calculated because ATC or ROA or Salt are not valid.");
                prevalid = false;
            }
            if (!IsMandatoryValidationValid(PACKSIZE_VALIDATION))
            {
                AddInfoMsg("Package content cannot be calculated because PACKSIZE is not valid.");
                prevalid = false;
            }
            if (!IsMandatoryValidationValid(STRENGTH_VALIDATION))
            {
                AddInfoMsg("Package content cannot be calculated because Strength or Concentration or Volume are not valid.");
                prevalid = false;
            }
            if (prevalid==false)
            {
                SetValidate(false, CONTENT_VALIDATION );
                return;
            }

            if (this.ATC5.Code == AMUConstants.ATC_Z99_CODE || (this.Combination?.Code == AMUConstants.COMB_Z99_CODE))
            {
                AddInfoMsg("Package content cannot be calculated because ATC5 or COMBINATIONS are not defined (Z99 codes).");
                return;
            }

            sAmount = new Amount
            { 
                Unit = this.StrengthUnit
            };
            if (this.Volume == null)
            {
                sAmount.Value = this.Strength;
            }
            else
            {
                sAmount.Value = this.Strength * (Decimal)this.Volume / (Decimal)this.ConcentrationVolume;
            }
            this.Content = new Amount(sAmount.Value * this.PackSize, sAmount.Unit);
            SetValidate(true, CONTENT_VALIDATION);
        }

        private void ValidateDddPerPackage()
        {
            Decimal stdDdd;
            Decimal stdContent;
            Decimal dpp;

            // Check if ARS is null, empty, or if DDD is not validated
            if (!IsMandatoryValidationValid(CONVERSION_VALIDATION) || !IsMandatoryValidationValid(CONTENT_VALIDATION) || !IsMandatoryValidationValid(DDD_VALIDATION))
            {
                SetValidate(false, DPP_VALIDATION);
                return;
            }

            if (this.ATC5.Code == AMUConstants.ATC_Z99_CODE || (this.Combination?.Code == AMUConstants.COMB_Z99_CODE))
            {
                SetValidate(false, DPP_VALIDATION);
                return;
            }

            
            // Convert base unit for DDD
            stdDdd = this.DDD.getStdAmountValue();
            stdContent = this.Content.getStdAmountValue();
            dpp = stdContent * this.ConversionFactor / stdDdd;

            this.NbDDD = dpp;
            SetValidate(true, DPP_VALIDATION);
        }

        private bool IsValidDataField(string fieldName)
        {
            return this.DataFieldValues[fieldName].IsValid;
        }

        private bool IsMandatoryValidationValid(string fieldName)
        {
            if (!this.Validations.ContainsKey(fieldName))
            {
                return false;
            }
            else
            {
                return Validations[fieldName];
            }
        }


        public void ValidateAware()
        {

            if (!IsMandatoryValidationValid(ARS_VALIDATION))
            {
                return;
            }
            // We exclude product without ATC5 code, Z99ZZ99 code and all not belonging to the ATB AM class
            if (this.AMClass != "ATB" || this.ATC5.Code == AMUConstants.ATC_Z99_CODE || this.ATC5 == null) 
            {
                this.AWaRe = AMUConstants.NA;
                this.SetValidate(true, AWR_VALIDATION);
                return;
            }
            
            var awr = ThisWorkbook.AwareDataList.Where(a => a.ATC5 == this.ATC5.Code && a.ROA == this.Roa.Code).FirstOrDefault();
            if (awr != null)
            {
                this.AWaRe = awr.AWR;
            }
            else // not classified in AWaRe, default to Not Classified
            {
                this.AWaRe = Aware.NotClassifiedCode;
            }
            SetValidate(true, AWR_VALIDATION);
        }

        public void ValidateMEML()
        {
            if (!IsMandatoryValidationValid(ARS_VALIDATION))
            {
                return;
            }

            if (this.ATC5 == null || this.ATC5.Code == AMUConstants.ATC_Z99_CODE)
            {
                this.MEML = YesNoNA.NA;
                SetValidate(true, MEML_VALIDATION);
                return;
            }

            var eml = ThisWorkbook.MemlDataList.Where(a => a.ATC5 == this.ATC5.Code && a.ROA == this.Roa.Code).FirstOrDefault();
            if (eml != null)
            {
                this.MEML= eml.EML;
            }
            else
            {
                this.MEML = YesNoNA.NA;
            }
            SetValidate(true, MEML_VALIDATION);
        }

        public bool GetValidate(string field)
        {
            this.Validations.TryGetValue(field, out bool ret);
            return ret;
        }

        public void SetValidate(bool val, string field)
        {
            if (Validations.ContainsKey(field))
            {
                Validations[field] = val;
            }
        }

        public object GetValueForVariable(string variable)
        {
            object value;

            switch (variable)
            {
                case UID_FIELD:
                    value = this.UniqueId;
                    break;

                case COUNTRY_FIELD:
                    value = Country;
                    break;

                case ENROLMENT_DATE_WIDP_FIELD:
                case INCIDENT_DATE_WIDP_FIELD:
                    value = new DateTime((int)Year, 1, 1);
                    break;

                case PRODUCT_ID_FIELD:
                    value = ProductId;
                    break;

                case PRODUCT_NAME_FIELD:
                    value = ProductName;
                    break;

                case LABEL_FIELD:
                    value = Label;
                    break;

                case PACKSIZE_FIELD:
                    value = PackSize;
                    break;

                case STRENGTH_FIELD:
                    value = Strength;
                    break;

                case STRENGTH_UNIT_FIELD:
                    value = StrengthUnit?.Code ?? null;
                    break;

                case CONCENTRATION_VOLUME_FIELD:
                    value = ConcentrationVolume;
                    break;

                case CONCENTRATION_VOLUME_UNIT_WIDP_FIELD:
                case VOLUME_UNIT_WIDP_FIELD:
                    value = "ML";
                    break;

                case VOLUME_FIELD:
                    value = Volume;
                    break;

                case ATC5_FIELD:
                    value = ATC5?.Code ?? null;
                    break;

                case COMBINATION_FIELD:
                    value = Combination?.Code ?? null;
                    break;

                case ROUTE_ADMIN_FIELD:
                    value = Roa?.Code ?? null;
                    break;

                case SALT_FIELD:
                    value = Salt?.Code ?? null;
                    break;

                case PAEDIATRIC_FIELD:
                    value = Paediatric!=null?YesNoUnknownString.GetStringFromYesNoUnk(Paediatric):null;
                    break;

                case FORM_FIELD:
                    value = Form;
                    break;

                case INGREDIENTS_FIELD:
                    value = Ingredients;
                    break;

                case PRODUCT_ORIGIN_FIELD:
                    value = ProductOrigin;
                    break;

                case MANUFACTURER_COUNTRY_FIELD:
                    value = ManufacturerCountry;
                    break;

                case MARKET_AUTH_HOLDER_FIELD:
                    value = MAHolder;
                    break;

                case GENERICS_FIELD:
                    value = Generics!=null ? YesNoUnknownString.GetStringFromYesNoUnk(Generics) : "";
                    break;

                case YEAR_AUTHORIZATION_FIELD:
                    value = YearAuthorization;
                    break;

                case YEAR_WITHDRAWAL_FIELD:
                    value = YearWithdrawal;
                    break;

                case DATA_STATUS_WIDP_FIELD:
                    value = 1;
                    break;

                default:
                    value = "";
                    break;
            }

            return value;
        }

    }
    #endregion
}
