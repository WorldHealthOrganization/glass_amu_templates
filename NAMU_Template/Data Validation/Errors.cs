// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.Collections.Generic;
using System.Linq;

namespace NAMU_Template.Data_Validation
{

    public enum ValidationMessageType
    {
        Info = 2,
        Warning = 3,
        Error = 4
    }
    public enum EntityStatus
    {
        OK = 1,
        INFO = 2,
        WARNING = 3,
        ERROR = 4,
        DEFAULT = -1
    }


    public class ErrorStatus
    {
        private List<string> errors = new List<string>();
        private List<string> infos = new List<string>();
        private EntityStatus status;
        private string errorType;

        public ErrorStatus()
        {
            errors = new List<string>();
            infos = new List<string>();
            status = EntityStatus.OK;
        }

        public string ErrorType
        {
            get { return errorType; }
            set { errorType = value; }
        }

        public List<string> Errors
        {
            get { return errors; }
        }

        public List<string> Infos
        {
            get { return infos; }
        }

        public EntityStatus Status
        {
            get { return status; }
            set { status = value; }
        }

        public void Reset()
        {
            errors = new List<string>();
            infos = new List<string>();
            status = EntityStatus.OK;
        }

        public void AddErrorMsgs(string msg)
        {
            errors.Add(msg);
            this.Status = EntityStatus.ERROR;
        }

        public void AddInfoMsg(string msg)
        {
            infos.Add(msg);
            if (this.Status == EntityStatus.OK)
            {
                this.Status = EntityStatus.INFO;
            }
        }

        public string ErrorsToString()
        {
            var elems = from er in errors select $"error: {er}";

            return String.Join("\n", elems);
        }

        public string InfosToString()
        {
            var elems = from info in infos select $"info: {info}";

            return String.Join("\n", elems);
        }
    }
}
