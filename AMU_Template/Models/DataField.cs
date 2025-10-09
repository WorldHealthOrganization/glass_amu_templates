// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;

namespace AMU_Template.Models
{

    public interface IDataField
    {
        public Type GenericType { get; set; }

        public object Value { get; set; }

        public string Name { get; set; }

        public bool IsValid { get; set; }

        public bool IsProvided { get;}

        public bool IsMissing { get; set; }

        public bool IsMandatory { get; set; }

        public int FieldColumn { get; set; }

        public string InputValue { get; set; }
    }

    public class DataField<T> : IDataField
    {

        public Type GenericType = typeof(T);

        public T Value { get; set; }

        public string Name { get; set; }

        public bool IsValid { get; set; } = false;

        public bool IsProvided {
            get {
                return !IsMissing && IsValid;
            }
        }

        public bool IsMandatory { get; set; } = false;

        public int FieldColumn { get; set; }

        object IDataField.Value
        {
            get { return this.Value; }
            set { this.Value = (T)value; }
        }

        public bool IsMissing { get; set; } = false;
        public string InputValue { get; set; }
        Type IDataField.GenericType { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
    }
}
