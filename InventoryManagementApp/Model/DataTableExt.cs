using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Collections;

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
                    where partNumList.Keys.Contains(item.Field<string>("PartNumber")) && so.Field<DateTime>("ShipDate") >= startDate && !so.Field<string>("Customer").Contains("Marlin Ste")
                    group so by new
                    {
                        Row = partNumList[item.Field<string>("PartNumber")].rowNum,
                        PartNumber = item.Field<string>("PartNumber"),
                        QtyOnHand = (int)item.Field<decimal>("QtyOnHand"),
                        RestockSODate = string.IsNullOrEmpty(partNumList[item.Field<string>("PartNumber")].restockSODate) ? 
                            Convert.ToString((DateTime?)null) : 
                            DateTime.FromOADate(Convert.ToDouble(partNumList[item.Field<string>("PartNumber")].restockSODate)).ToShortDateString()
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
                        AvgSalePrice = String.Format("{0:C}", itemGroup.Average(so => so.Field<decimal>("SalePrice"))),
                        Last15Months = (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity"))),
                        MaxStockRev = (int)((itemGroup.Average(so => so.Field<decimal>("SalePrice")) * (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 3m / 15m))) > 1000 ?
                            String.Format("{0:C}", (int)((itemGroup.Average(so => so.Field<decimal>("SalePrice")) * (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 3m / 15m)))) :
                            String.Format("{0:C}", (int)((1000m / itemGroup.Average(so => so.Field<decimal>("SalePrice"))) * (itemGroup.Average(so => so.Field<decimal>("SalePrice"))))),
                        RestockSODate = itemGroup.Key.QtyOnHand >= (int)(itemGroup.Sum(so => so.Field<decimal>("Quantity")) * 1.5m / 15.0m) ? "" : itemGroup.Key.RestockSODate
                    };

            dt = minMaxGroup.CustomCopyToDataTable();
            dt.PrimaryKey = new DataColumn[] { dt.Columns["PartNumber"] };

            return dt;
        }

        public static DataTable BuildSOReqTable(this DataTable soReqTable, DataTable minMaxDt)
        {
            DataTable dt = new DataTable();

            bool isDuringQuarterlyRush = IsDuringQuarterlyRush();

            try
            {
                IEnumerable<DataRow> minMaxRows =
                    from item in minMaxDt.AsEnumerable()
                    where (item.Field<int>("QtyOnHand") < item.Field<int>("Min") || (isDuringQuarterlyRush ? item.Field<int>("QtyOnHand") <= item.Field<int>("Min") * 1.1 : false)) && string.IsNullOrEmpty(item.Field<string>("RestockSODate"))
                    orderby item["Row"] ascending
                    select item;

                dt = minMaxRows.CopyToDataTable<DataRow>();
                dt.PrimaryKey = new DataColumn[] { dt.Columns["PartNumber"] };
            }
            catch { }
                        
            return dt;
        }

        public static DataTable BuildPending(this DataTable pendingDt, DataTable minMaxDt)
        {
            DataTable dt = new DataTable();         
            try
            {
                IEnumerable<DataRow> minMaxRows =
                    from item in minMaxDt.AsEnumerable()
                    where (item.Field<int>("QtyOnHand") < item.Field<int>("Min")) && !string.IsNullOrEmpty(item.Field<string>("RestockSODate"))
                    orderby item["Row"] ascending
                    select item;

                dt = minMaxRows.CopyToDataTable<DataRow>();
                dt.PrimaryKey = new DataColumn[] { dt.Columns["PartNumber"] };
            }
            catch { }

            return dt;
        }

        private static bool IsDuringQuarterlyRush()
        {
            
            DateTime currentDate = new DateTime(2018, DateTime.Today.Month, DateTime.Today.Day);
            System.Diagnostics.Debug.WriteLine(currentDate.ToString());
                                    
            List<DateRange> dateRanges = new List<DateRange>()
            {
                new DateRange(new DateTime(2018, 2, 20), new DateTime(2018, 3, 31)),
                new DateRange(new DateTime(2018, 5, 20), new DateTime(2018, 6, 30)),
                new DateRange(new DateTime(2018, 8, 20), new DateTime(2018, 9, 30)),
                new DateRange(new DateTime(2018, 11, 20), new DateTime(2018, 12, 31))
            };

            foreach(DateRange dateRange in dateRanges)
            {
                if (currentDate >= dateRange.startDate && currentDate <= dateRange.endDate)
                {
                    return true;
                }
            }

            return false;
            
        }

        class DateRange
        {
            public DateTime startDate;
            public DateTime endDate;

            public DateRange(DateTime startDate, DateTime endDate)
            {
                this.startDate = startDate;
                this.endDate = endDate;
            }
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
