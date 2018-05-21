using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagementApp
{
    /// <summary>
    /// Struct that represents the mapping of the columns in Excel Inventory Management Doc
    /// </summary>
    public static class ExcelColumn
    {
        public readonly static string min = "B";
        public readonly static string max = "C";
        public readonly static string onHand = "D";
        public readonly static string avgSalePrice = "E";
        public readonly static string quantitySold = "F";
        public readonly static string maxStockRev = "G";
    }
}
