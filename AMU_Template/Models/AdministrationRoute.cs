// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace AMU_Template.Models
{
    public class AdministrationRoute
    {
        
        public string Code { get; set; }

        public string Name { get; set; }

        public AdministrationRoute()
        {

        }

        public AdministrationRoute(string code, string name)
        {
            this.Code = code;
            this.Name = name;
        }
    }
}
