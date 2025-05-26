// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Constants;

namespace AMU_Template.Models
{
    public class MEML
    {

        public string ATC5 { get; set; }

        public string ROA { get; set; }

        public YesNoNA EML { get; set; }

        public MEML() { }   

        public MEML(string atc5, string roa, YesNoNA inEml)
        {
            ATC5 = atc5;
            ROA = roa;
            EML = inEml;
        }
    }
}
