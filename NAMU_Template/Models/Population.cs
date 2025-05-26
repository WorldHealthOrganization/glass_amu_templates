// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using NAMU_Template.Constants;

namespace NAMU_Template.Models
{

    public class PopulationKey
    {
        public int Year { get; }
        public string AMClass { get; }
        public HealthSector Sector { get; }

        private int hashCode;

        public PopulationKey(int year, string amClass, HealthSector sector)
        {
            Year = year;
            AMClass = amClass;
            Sector = sector;
            hashCode = (Year, AMClass, Sector).GetHashCode();
        }

        public override bool Equals(object obj)
        {
            return obj is PopulationKey key &&
                   Year == key.Year &&
                   AMClass == key.AMClass &&
                   Sector == key.Sector;
        }

        public override int GetHashCode()
        {
            return hashCode;
        }
    }

    public class Population
    {
        public string Country { get; set; }

        public int Year { get; set; }

        public HealthSector Sector { get; set; }

        public string ATCClass { get; set; }

        public Decimal? TotalPopulation { get; set; }

        public Decimal? CommunityPopulation { get; set; }

        public Decimal? HospitalPopulation { get; set; }

        // Computed key for fast lookup
        public PopulationKey Key => new PopulationKey(Year, ATCClass, Sector);
    }
}
