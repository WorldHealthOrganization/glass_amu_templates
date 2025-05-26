// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using AMU_Template.Helpers;

namespace AMU_Template.Models
{
    public class DDD
    {
        public string ARS { get; set; }

        public ATC ATC5 { get; set; }

        public AdministrationRoute ROA { get; set; }

        public Salt Salt { get; set; }

        public Decimal Value { get; set; }

        public MeasureUnit Unit { get; set; }

        public Decimal StdValue { get; set; }

        // public double BaseDDD { get; set; }

        public string? Notes { get; set; }

        public DDD(ATC ATC5, AdministrationRoute ROA, Salt Salt, Decimal Value, MeasureUnit Unit, Decimal StdValue, string? Notes)
        {
            this.ATC5 = ATC5;
            this.ROA = ROA;
            this.Salt = Salt;
            this.Value = Value;
            this.Unit = Unit;
            this.StdValue = StdValue;
            this.Notes = Notes;
            this.ARS = ARSHelper.GenerateARSFromATC5ROASalt(this.ATC5.Code, this.ROA.Code, this.Salt.Code);
        }
    }
}
