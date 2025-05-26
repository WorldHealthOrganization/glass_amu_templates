// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace AMU_Template.Models
{
    public class Aware
    {
        public const string NotClassifiedCode = "N";
        public string ATC5 { get; set; }

        public string ROA { get; set; }

        public string AWR { get; set; }

        public Aware() { }

        public Aware(string atc5, string roa, string awr)
        {
            ATC5 = atc5;
            ROA = roa;
            AWR = awr;
        }
    }
}
