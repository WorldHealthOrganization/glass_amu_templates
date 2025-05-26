// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using NAMU_Template.Constants;

namespace NAMU_Template.Models
{

    public class DataAvailabilityKey
    {
        public int Year { get; }
        public string AMClass { get; }
        public HealthSector Sector { get; }

        private int hashCode;

        public DataAvailabilityKey(int year, string amClass, HealthSector sector)
        {
            Year = year;
            AMClass = amClass;
            Sector = sector;
            hashCode = (Year, AMClass, Sector).GetHashCode();
        }

        public override bool Equals(object obj)
        {
            return obj is DataAvailabilityKey key &&
                   Year == key.Year &&
                   AMClass == key.AMClass &&
                   Sector == key.Sector;
        }

        public override int GetHashCode()
        {
            return hashCode;
        }
    }

    public class DataAvailability
    {
        public string Country { get; set; }

        public int Year { get; set; }

        public string ATCClass { get; set; }

        public HealthSector Sector { get; set; }

        public bool AvailabilityTotal { get; set; }

        public bool AvailabilityCommunity { get; set; }

        public bool AvailabilityHospital { get; set; }


        public List<int> Years;
        public class Availability
        {
            public bool Total;

            public bool Community;

            public bool Hospital;

            public Availability()
            {
                Total = false;
                Community = false;
                Hospital = false;
            }

            public bool IsAvailable()
            {
                return Total || Community || Hospital;
            }
        }
        // **Precomputed Key for Fast Lookup**
        public DataAvailabilityKey Key => new DataAvailabilityKey(Year, ATCClass, Sector);
    }
}
