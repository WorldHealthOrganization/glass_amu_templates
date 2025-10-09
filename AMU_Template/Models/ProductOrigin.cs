// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace AMU_Template.Models
{
    public class ProductOrigin
    {
        public string Code { get; set; }
        public string Name { get; set; }


        public ProductOrigin() { }

        public ProductOrigin(string code, string name)
        {
            Code = code;
            Name = name;
        }
    }
}
