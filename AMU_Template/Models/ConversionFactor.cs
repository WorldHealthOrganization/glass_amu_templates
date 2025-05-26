// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;

namespace AMU_Template.Models
{
    public class ConversionFactor
    {
        public string ARS { get; set; }

        public ATC ATC5 { get; set; }

        public AdministrationRoute ROA { get; set; }

        public Salt Salt { get; set; }

        public MeasureUnit UnitFrom { get; set; }

        public MeasureUnit UnitTo { get; set; }

        public Decimal Factor { get; set; }
        
    }
}
