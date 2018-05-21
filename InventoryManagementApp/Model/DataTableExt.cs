using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace InventoryManagementApp
{
    static class DataTableExt
    {
        public static DataTable BuildTable(this DataTable minMaxDt, QuickBooksDataTable salesOrderDataTable, QuickBooksDataTable itemDataTable, Dictionary<string, int> partNumList)
        {
            DataTable dt = new DataTable();
            DateTime startDate = DateTime.Today.AddMonths(-15); // Rolling 15 months

            var minMaxGroup =
                    from item in itemDataTable.AsEnumerable()
                    join so in salesOrderDataTable.AsEnumerable()
                    on item.Field<string>("PartNumber") equals so.Field<string>("PartNumber")
                    where partNumList.Keys.Contains(item.Field<string>("PartNumber")) && so.Field<DateTime>("ShipDate") >= startDate
                    group so by new
                    {
                        PartNumber = item.Field<string>("PartNumber"),
                        QtyOnHand = item.Field<decimal>("QtyOnHand")
                    } into itemGroup
                    select new
                    {
                        PartNumber = itemGroup.Key.PartNumber,
                        QtyOnHand = itemGroup.Key.QtyOnHand,
                        Min = (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 1.5m / 15.0m),
                        Max = (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 3m / 15m),
                        Last15Months = (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity"))),
                        AvgSalePrice = itemGroup.Average(so => so.Field<decimal>("SalePrice")),
                        MaxStockRev = (int)(itemGroup.Average(so => so.Field<decimal>("SalePrice")) * (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 3m / 15m)),
                    };

            dt = minMaxGroup.CopyToDataTable();
            dt.PrimaryKey = new DataColumn[] { dt.Columns["PartNumber"] };

            return dt;
        }

        public static void Write(this DataTable dataTable, string filepath)
        {
            // Read table, while there is still a record
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(filepath))
            {
                StringBuilder sb = new StringBuilder();

                // Write Column Headers to TXT
                IEnumerable<string> columnHeaders = dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                sb.AppendLine(string.Join(",", columnHeaders));

                Console.WriteLine("Data Header Written.");

                foreach (DataRow row in dataTable.Rows)
                {
                    try
                    {
                        IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                        sb.AppendLine(string.Join(",", fields));
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Debug.WriteLine(e.Message);
                        Console.WriteLine("Failed Line Write.");
                    }
                }

                writer.Write(sb.ToString());
                Console.WriteLine("Data Written to CSV File.");
            }
        }   
    }    
}
