// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using NAMU_Template.Data_Validation;

namespace NAMU_Template.Models
{
    public class ValidationMessage
    {
        public ValidationMessageType MessageType;
        public string Message;
        public object ErrorField; //This will store the Column value of the error field..!
    }
}
