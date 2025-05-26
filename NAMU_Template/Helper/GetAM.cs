// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;

namespace NAMU_Template.Helper
{
    public class GetAM
    {
        public const string A07AA_CLASS = "A07AA";
        public const string D01BA_CLASS = "D01BA";
        public const string P01AB_CLASS = "P01AB";
        public const string P01B_CLASS = "P01B";
        public const string J01_CLASS = "J01";
        public const string J02_CLASS = "J02";
        public const string J04_CLASS = "J04";
        public const string J05_CLASS = "J05";

        public static string GetAMClassForATC(string atc5)
        {
            // Extract the first 5 characters of the string 
            string code = atc5.Substring(0, Math.Min(atc5.Length, 5));

            if (code == A07AA_CLASS)
            {
                return A07AA_CLASS;
            }
            if (code == D01BA_CLASS)
            {
                return D01BA_CLASS;
            }
            if (code == P01AB_CLASS)
            {
                return P01AB_CLASS;
            }
            if (code == P01B_CLASS)
            {
                return P01B_CLASS;
            }
            if (code == J01_CLASS)
            {
                return J01_CLASS;
            }
            if (code == J02_CLASS)
            {
                return J02_CLASS;
            }
            if (code == J04_CLASS)
            {
                return J04_CLASS;
            }
            if (code == J05_CLASS)
            {
                return J05_CLASS;
            }

            // Extract the first 3 charc of the string 

            code = atc5.Substring(0, Math.Min(atc5.Length, 3));

            if (code == J01_CLASS)
            {
                return J01_CLASS;
            }

            if (code == J02_CLASS)
            {
                return J02_CLASS;
            }

            if (code == J04_CLASS)
            {
                return J04_CLASS;
            }

            if (code == J05_CLASS)
            {
                return J05_CLASS;
            }
            if (code == P01AB_CLASS)
            {
                return P01AB_CLASS;
            }
            if (code == P01B_CLASS)
            {
                return P01B_CLASS;
            }

            if (code == A07AA_CLASS)
            {
                return A07AA_CLASS;
            }
            if (code == D01BA_CLASS)
            {
                return D01BA_CLASS;
            }

            code = atc5.Substring(0, Math.Min(atc5.Length, 4));
            if (code == P01B_CLASS)
            {
                return P01B_CLASS;
            }


            // Default return if no match is found
            return null;

        }
    }
}
