// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace AMU_Template.Models
{
    public class ATC
    {
  
        public string Code { get; set; }

        public string Name { get; set; }

        public int Level { get; set; }

        private string InternalParentCode;

        public string? ParentCode { 
            get
            {
                if(InternalParentCode != null)
                {
                    return InternalParentCode;
                }
                else
                {
                    int n = 0;
                    switch (Level)
                    {
                        case 2:
                            n = 1;
                            break;
                        case 3:
                            n = 3;
                            break;
                        case 4:
                            n = 4;
                            break;
                        case 5:
                            n = 5;
                            break;
                    }
                    if (n != 0)
                    {
                        this.InternalParentCode = this.Code.Substring(0, n);
                        return this.InternalParentCode;
                    }
                    else { return null; }
                }
            }
        }

        public string AMClass { get; set; }

        public string ATCClass { get; set; }

    }
}
