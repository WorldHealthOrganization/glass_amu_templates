// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

namespace NAMU_Template.Models
{
    public class WidpVariable
    {
        private string _code;
        private string _label;
        private int _index;

        // Property for code
        public string Code
        {
            get { return _code; }
            private set { _code = value; }
        }

        // Property for name (assuming name corresponds to label in VBA)
        public string Label
        {
            get { return _label; }
            private set { _label = value; }
        }

        // Property for index
        public int Index
        {
            get { return _index; }
            private set { _index = value; }
        }

        public WidpVariable(string code, string label, int index)
        {
            Code = code;
            Label = label;
            Index = index;
        }
    }
}
