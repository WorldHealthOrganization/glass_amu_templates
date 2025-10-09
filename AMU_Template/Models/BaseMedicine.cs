// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;

namespace AMU_Template.Models
{

    public class Amount
    {
        public Decimal Value { get; set; }
        public MeasureUnit Unit { get; set; }


        public Amount()
        {
            
        }

        public Amount(decimal value, MeasureUnit unit)
        {
            Value = value;
            Unit = unit;
        }

        public Decimal getStdAmountValue()
        {
            return this.Value * this.Unit.BaseConversion;
        }
    }

    public abstract class BaseMedicine
    {
        public abstract string UniqueId { get; }

        public string Label { get; set; }

        public ATC ATC5 { get; set; }

        public AdministrationRoute Roa { get; set; }

        public string AMClass { get; set; }

        public string ATCClass { get; set; }

        public string ARS { get; set; }

        public Amount? Content { get; set; }

        public Amount? DDD { get; set; }

        public Decimal NbDDD { get; set; }
    }
}
