using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace InventoryManagementApp.Model
{
    static class DataTableExt
    {
        public static DataTable BuildTable(this DataTable minMaxDt, IQuickBooksData salesOrderDataTable, IQuickBooksData itemDataTable, Dictionary<string, ExcelPartNumber> partNumList)
        {
            DataTable dt = new DataTable();
            DateTime startDate = DateTime.Today.AddMonths(-15); // Rolling 15 months

            var minMaxGroup =
                    from item in ((QuickBooksDataTable)itemDataTable).AsEnumerable()
                    join so in ((QuickBooksDataTable)salesOrderDataTable).AsEnumerable()
                    on item.Field<string>("PartNumber") equals so.Field<string>("PartNumber")
                    where partNumList.Keys.Contains(item.Field<string>("PartNumber")) && so.Field<DateTime>("ShipDate") >= startDate
                    group so by new
                    {
                        Row = partNumList[item.Field<string>("PartNumber")].rowNum,
                        PartNumber = item.Field<string>("PartNumber"),
                        QtyOnHand = (int)item.Field<decimal>("QtyOnHand"),
                        RestockSONum = partNumList[item.Field<string>("PartNumber")].restockSONum,
                        RestockSODate = partNumList[item.Field<string>("PartNumber")].restockSODate
                    } into itemGroup
                    select new
                    {
                        Row = itemGroup.Key.Row + 1,
                        PartNumber = itemGroup.Key.PartNumber,
                        Min = (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 1.5m / 15.0m),
                        Max = (int)((itemGroup.Average(so => so.Field<decimal>("SalePrice")) * (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 3m / 15m))) > 1000 ? 
                            (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 3m / 15m) : 
                            (int)(1000m / itemGroup.Average(so => so.Field<decimal>("SalePrice"))),
                        QtyOnHand = itemGroup.Key.QtyOnHand,
                        AvgSalePrice = Math.Round(itemGroup.Average(so => so.Field<decimal>("SalePrice")), 2),
                        Last15Months = (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity"))),
                        MaxStockRev = (int)((itemGroup.Average(so => so.Field<decimal>("SalePrice")) * (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 3m / 15m))) > 1000 ?
                            (int)((itemGroup.Average(so => so.Field<decimal>("SalePrice")) * (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 3m / 15m))) :
                            (int)((1000m / itemGroup.Average(so => so.Field<decimal>("SalePrice"))) * (itemGroup.Average(so => so.Field<decimal>("SalePrice")))),
                        RestockSONum = itemGroup.Key.RestockSONum,
                        RestockSODate = itemGroup.Key.RestockSODate
                    };

            dt = minMaxGroup.CustomCopyToDataTable();
            dt.PrimaryKey = new DataColumn[] { dt.Columns["PartNumber"] };

            return dt;
        }

        public static DataTable BuildSOReqTable(this DataTable soReqTable, DataTable minMaxDt)
        {
            DataTable dt = new DataTable();
 
            IEnumerable<DataRow> minMaxRows =
                    from item in minMaxDt.AsEnumerable()
                    where (item.Field<int>("QtyOnHand") < item.Field<int>("Min")) && string.IsNullOrEmpty(item.Field<string>("RestockSONum"))
                    orderby item["Row"] ascending
                    select item;

            dt = minMaxRows.CopyToDataTable<DataRow>();
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
