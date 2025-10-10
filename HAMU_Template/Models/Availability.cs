using HAMU_Template.Constants;

namespace HAMU_Template.Models
{

    public class AvailabilityKey
    {
        private int hashCode;
        public string details {get;}

        public AvailabilityKey(string country, int year, string hospital, FacilityStructureLevel level, string atcClass)
        {
            hashCode = (country, year, hospital, level, atcClass).GetHashCode();
            details = $"{country}|{year}|{hospital}|{FacilityStructureLevelString.GetStringForFacilityStructureLevel(level)}|{atcClass}";
        }

        public override bool Equals(object obj)
        {
            return obj is AvailabilityKey key &&
                hashCode == key.hashCode &&
                details == key.details;    
        }

        public override int GetHashCode()
        {
            return hashCode;
        }
    }

    public class Availability
    {
        public string Country { get; set; }

        public int Year { get; set; }

        public string ATCClass { get; set; }

        public FacilityStructureLevel Level { get; set; }

        public string Hospital { get; set; }

        // **Precomputed Key for Fast Lookup**
        public AvailabilityKey Key => new AvailabilityKey(Country, Year, Hospital, Level, ATCClass);
    }


    //public class AvailabilityData
    //{
    //    public string Country { get; set; }

    //    public string Hospital { get; set; }

    //    public int Year { get; set; }

    //    public bool A07AA_Class {  get; set; }

    //    public bool D01BA_Class { get; set; }
    
    //    public bool J01_Class { get; set; }

    //    public bool J02_Class { get; set; } 

    //    public bool J04_Class { get; set; }

    //    public bool J05_Class { get; set; }

    //    public bool P01AB_Class { get; set; }

    //    public bool P01B_Class { get; set; }

    //    public FacilityStructureLevel Level { get; set; }
    //}
}
