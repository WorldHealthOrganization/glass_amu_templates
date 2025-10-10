using AMU_Template.Constants;
using HAMU_Template.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HAMU_Template.Models
{

    public abstract class MedicineConsumption
    {

        // WIDP Excel Template Fields
        public const string MED_CONS_UID_FIELD = "MED_CONS_UID_FIELD";
        public const string MED_CONS_COUNTRY_FIELD = "MED_CONS_UID_FIELD";

        public const string PROD_CONS_H_SECTOR_FIELD = "PROD_CONS_H_SECTOR_FIELD";
        public const string PROD_CONS_H_LEVEL_FIELD = "PROD_CONS_H_LEVEL_FIELD";
        public const string PROD_CONS_PACKAGES_FIELD = "PROD_CONS_PACKAGES_FIELD";

        public static Decimal Hundred = new Decimal(100);
        

        public string Key { get; set; }

        public int LineNo { get; set; }

        public int Sequence { get; set; }

        public string Country { get; set; }

        public string Hospital {  get; set; }

        public FacilityStructureLevel Level {  get; set; }

        public string Structure { get; set; }

        public int Year { get; set; }

        public string AMClass { get; set; }

        public string AtcClass { get; set; }

        public string AWaRe { get; set; }

        public YesNoNA MEML { get; set; }

        public string UniqueId { get; set; }

        public string Label { get; set; }

        public string ATC5 { get; set; }

        protected string InternalATC4;

        public string ATC4
        {
            get
            {
                if (!string.IsNullOrEmpty(InternalATC4)) { return InternalATC4; }
                else
                {
                    if (!string.IsNullOrEmpty(ATC5))
                    {
                        InternalATC4 = ATC5.Substring(0, 5);
                        return InternalATC4;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }

        protected string InternalATC3;
        public string ATC3
        {
            get
            {
                if (!string.IsNullOrEmpty(InternalATC3)) { return InternalATC3; }
                else
                {
                    if (!string.IsNullOrEmpty(ATC5))
                    {
                        InternalATC3 = ATC5.Substring(0, 4);
                        return InternalATC3;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }

        protected string InternalATC2;
        public string ATC2
        {
            get
            {
                if (!string.IsNullOrEmpty(InternalATC2)) { return InternalATC2; }
                else
                {
                    if (!string.IsNullOrEmpty(ATC5))
                    {
                        InternalATC2 = ATC5.Substring(0, 3);
                        return InternalATC2;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
        }

        public string Roa { get; set; }

        public YesNoUnknown Paediatric { get; set; }

        public Decimal? BedDays { get; set; }
        
        public Decimal? Admissions { get; set; }

        public Decimal DDD { get; set; } = Decimal.Zero;
        public Decimal DBD { get; set; } = Decimal.Zero;
        public Decimal DAD { get; set; } = Decimal.Zero;


        public abstract void CalculateDDD();


        public abstract Decimal GetPackUnit();


        public void CalculateDDDPerActivity()
        {
            if (DDD > Decimal.Zero)
            {
                if (Admissions != Decimal.Zero)
                {
                    DAD = (decimal)(DDD * Hundred / Admissions);
                }
                if (BedDays != Decimal.Zero)
                {
                    DBD = (decimal)(DDD * Hundred / BedDays);
                }
            }
        }
    }


    public class ProductConsumption: MedicineConsumption
    {
        public string ProductId { get; set; }

        public Decimal Packages { get; set; }

        public Decimal DPP { get; set; }


        public override void CalculateDDD()
        {
            if (DPP > 0)
            {
                DDD = Packages * DPP;
            }
        }

        public override decimal GetPackUnit()
        {
            return Packages;
        }
    }

    public class SubstanceConsumption : MedicineConsumption
    {
 
        public Decimal Units { get; set; }

        public Decimal DPP { get; set; }

        public string Structure { get; set; }

        public FacilityStructureLevel Level { get; set; }


        public override void CalculateDDD()
        {
            if (DPP > 0)
            {
                DDD = Units * DPP;
            }
        }

        public override decimal GetPackUnit()
        {
            return Units;
        }
    }
}
