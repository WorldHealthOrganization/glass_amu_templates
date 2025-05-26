// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System.Collections.Generic;
using NAMU_Template.Models;

namespace NAMU_Template.Helper
{
    public class SharedData
    {
        public static List<DataAvailability> AvailData { get; set; } = new List<DataAvailability>();
        public static List<Population> PopYears { get; set; } = new List<Population>();

        // A static property to hold the products dictionary
        public static List<Product> Products { get; set; } = new List<Product>();
        public static List<ProductConsumption> ProductConsummptionData { get; set; } = new List<ProductConsumption>();

        // Holds DDD consumption data
        public static List<AtcConsumption> AtcConsumptionData { get; set; } = new List<AtcConsumption>();


    }
}
