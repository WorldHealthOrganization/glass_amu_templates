// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;

namespace AMU_Template.Models
{
    public class DDDCombination
    {

        public string Code{ get; set; }

        public ATC ATC5 { get; set; }

        public AdministrationRoute ROA { get; set; }

        public string Form {  get; set; }

        public string UnitDose { get; set; }

        public Decimal DDDValue { get; set; }

        public MeasureUnit DDDUnit { get; set; }

        public string? Info { get; set; }

        public string? Examples { get; set; }

    }
}
