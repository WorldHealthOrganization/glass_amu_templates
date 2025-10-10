// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;
using AMU_Template.Constants;
using AMU_Template.Helpers;
using AMU_Template.Models;
using AMU_Template.Validations;

namespace HAMU_Template.Models
{
    public class Product : Medicine
    {

        // WHO Excel Template Fields

        public const string PRODUCT_ID_FIELD = "PRODUCT_ID";
        public const string PRODUCT_NAME_FIELD = "PRODUCT_NAME";
        public const string PRODUCT_ORIGIN_FIELD = "PRODUCT_ORIGIN";
        public const string GENERICS_FIELD = "GENERICS";

        
        // Validation constants

        public const string PRODUCT_ID_VALIDATION = "PRODUCT_ID_VALIDATION";
        public const string PACKSIZE_VALIDATION = "PACKSIZE_VALIDATION";

        public string ProductId { get; set; }

        public string ProductName { get; set; }

        public string ProductOrigin { get; set; }

        public YesNoUnknown Generics { get; set; }

        public override string UniqueId
        {
            get
            { if (InternalCachedProductUniqueId != null)
                {
                    return InternalCachedProductUniqueId;
                }
                else
                {
                    if (LineNo == 0 || string.IsNullOrEmpty(ProductId))
                    {
                        return "UNDEFINED";
                    }
                    else
                    {
                        InternalCachedProductUniqueId = $"{LineNo}|{ProductId}";
                        return InternalCachedProductUniqueId;
                    }
                }
            }
        }

        public Product(): base()
        {
            DataFieldValues.Add(PRODUCT_ID_FIELD, null);
            DataFieldValues.Add(PACKSIZE_FIELD, null);
            DataFieldValues.Add(PRODUCT_NAME_FIELD, null);
            DataFieldValues.Add(PRODUCT_ORIGIN_FIELD, null);
            DataFieldValues.Add(GENERICS_FIELD, null);
            
            Validations.Add(PRODUCT_ID_VALIDATION, false);
            Validations.Add(PACKSIZE_VALIDATION, false);
        }

        #region Validation

        public override void Validate(bool force)
        {
            ValidateInit(force);
            ValidateOptionalFields();
            ValidateMedicine();
            ValidateDerived(force);
            FinalizeValidation();
        }

        protected override void ValidateDerived(bool force)
        {
            ValidateProductId();
            ValidatePackSize();
            ValidatePackageContent();
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

        private new void ValidatePackageContent()
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


       
        public override object GetValueForVariable(string variable)
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

                case GENERICS_FIELD:
                    value = Generics!=null ? YesNoUnknownString.GetStringFromYesNoUnk(Generics) : "";
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
