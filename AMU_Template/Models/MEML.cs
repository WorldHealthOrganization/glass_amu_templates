// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Constants;

namespace AMU_Template.Models
{
    public class MEML
    {

        public string ATC5 { get; set; }

        public string ROA { get; set; }

        public string INN { get; set; }

        public string Equivalent { get; set; }

        public MEML() { }   

        public MEML(string atc5, string roa, string inn, string equivalent)
        {
            ATC5 = atc5;
            ROA = roa;
            INN = inn;
            Equivalent = equivalent;
        }
    }
}
