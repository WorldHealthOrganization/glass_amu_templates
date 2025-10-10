using HAMU_Template.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HAMU_Template.Models.Mappings
{

    public class HospitalActivityKey
    {
        private int hashCode;
        public string details { get; }

        public HospitalActivityKey(string country, int year, string hospital, FacilityStructureLevel level, string structure)
        {
            hashCode = (country, year, hospital, FacilityStructureLevelString.GetStringForFacilityStructureLevel(level), structure).GetHashCode();
            details = $"{country}|{year}|{hospital}|{FacilityStructureLevelString.GetStringForFacilityStructureLevel(level)}|{structure}";
        }

        public override bool Equals(object obj)
        {
            return obj is HospitalActivityKey key &&
                hashCode == key.hashCode &&
                details == key.details;
        }

        public override int GetHashCode()
        {
            return hashCode;
        }
    }

    public class HospitalActivity
    {

        public string Country { get; set; }

        public string Hospital { get; set; }

        public int Year { get; set; }

        public FacilityStructureLevel Level { get; set; }

        public string Structure { get; set; }

        public int PatientDays { get; set; }

        public int Admissions { get; set; }

        public HospitalActivityKey Key => new HospitalActivityKey(Country, Year, Hospital, Level, Structure);

    }
}
