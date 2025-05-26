// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;

namespace AMU_Template.Models
{
    public class Salt
    {
        public string Code { get; set; }

        public string Name { get; set; }

        public string? Info { get; set; }

        public List<string> Atc5s { get; set; } = new List<string>();


        public Salt(string code, string name, string? info) {
            Code = code;
            Name = name;
            Info = info;
        }

        public Salt(string code, string name, string? info, List<String> atc5s)
        {
            Code = code;
            Name = name;
            Info = info;
            Atc5s = atc5s;
        }
    }
}
