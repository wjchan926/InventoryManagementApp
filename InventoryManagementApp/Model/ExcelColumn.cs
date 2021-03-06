﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagementApp.Model
{
    /// <summary>
    /// Class that represents the mapping of the columns in Excel Inventory Min-Max Management Doc.
    /// </summary>
    public static class ExcelColumn
    {
        public readonly static string partNumber = "A";
        public readonly static string min = "B";
        public readonly static string max = "C";
        public readonly static string onHand = "D";
        public readonly static string avgSalePrice = "E";
        public readonly static string quantitySold = "F";
        public readonly static string maxStockRev = "G";
        public readonly static string restockSODate = "H";
        public readonly static string bracketsPerSheet = "I";
    }
}
