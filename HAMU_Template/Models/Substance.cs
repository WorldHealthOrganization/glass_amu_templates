// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Constants;
using AMU_Template.Helpers;
using AMU_Template.Models;
using AMU_Template.Validations;
using System;
using System.Collections.Generic;
using System.Linq;
using static System.Collections.Specialized.BitVector32;

namespace HAMU_Template.Models
{
    public class Substance : Medicine
    {

        public override void Validate(bool force)
        {
            ValidateInit(force);
            ValidateOptionalFields();
            ValidateMedicine();
            ValidateDerived(force);
            FinalizeValidation();
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
                    value = Paediatric != null ? YesNoUnknownString.GetStringFromYesNoUnk(Paediatric) : null;
                    break;

                case FORM_FIELD:
                    value = Form;
                    break;

                case INGREDIENTS_FIELD:
                    value = Ingredients;
                    break;


                default:
                    value = null;
                    break;
            }

            return value;
        }

    }

}