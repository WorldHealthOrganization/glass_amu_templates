// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NAMU_Template.Models
{
    public class Country
    {
        public string Code { get; set; }

        public string ShortName { get; set; }

        public string FormalName { get; set; }


        public Country() { }

        public Country(string code, string shortName, string formalName)
        {
            Code = code;
            ShortName = shortName;
            FormalName = formalName;
        }
    }
}
