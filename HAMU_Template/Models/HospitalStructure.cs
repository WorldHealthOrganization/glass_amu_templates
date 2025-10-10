using HAMU_Template.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HAMU_Template.Models
{

    public class HospitalStructureKey
    {
        private int hashCode;
        public string details { get; }

        public HospitalStructureKey(string country, int year, string hospital)
        {
            hashCode = (country, year, hospital).GetHashCode();
            details = $"{country}|{year}|{hospital}";
        }

        public override bool Equals(object obj)
        {
            return obj is HospitalStructureKey key &&
                hashCode == key.hashCode &&
                details == key.details;
        }

        public override int GetHashCode()
        {
            return hashCode;
        }
    }

    public class HospitalStructure
    {
        public string Country { get; set; }
        public string Hospital { get; set; }
        public int? Year { get; set; }

        public HospitalStructureKey Key => new HospitalStructureKey(Country, (int)Year, Hospital);
    }
}
