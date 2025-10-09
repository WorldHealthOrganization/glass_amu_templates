// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using AMU_Template.Validations;

namespace AMU_Template.Validations
{
    public class ValidationMessage
    {
        public ValidationMessageType MessageType;
        public string Message;
        public object ErrorField; //This will store the Column value of the error field..!
    }
}
