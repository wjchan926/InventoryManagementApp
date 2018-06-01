using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Collections;

namespace InventoryManagementApp.Model
{
    /// <summary>
    /// Contains all DataTable extension methods.
    /// </summary>
    static class DataTableExt
    {
        /// <summary>
        /// Builds the MinMax DataTable.  This DataTable Holds all the information that will be written into the Min-Max document.
        /// </summary>
        /// <param name="minMaxDt">Current DataTable where information queried from QB will be stored.</param>
        /// <param name="salesOrderDataTable">DataTable that contains Sales Orders information from QB.</param>
        /// <param name="itemDataTable">DataTable that contains ItemAssembly information from QB.</param>
        /// <param name="partNumList">Part Numbers from the Min-Max document for querying.</param>
        /// <returns>A DataTable that represents the Min-Max Data to be written onto the Min-Max Document.</returns>
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

        /// <summary>
        /// Builds the DataTable that shows which parts need Restock Sales Orders made.  Uses the MinMax DataTable.
        /// </summary>
        /// <param name="soReqTable">Current DataTable where information will be stored.</param>
        /// <param name="minMaxDt">Min-Max DataTable.</param>
        /// <returns>A DataTable of Part Numbers where Restock Sales Orders need to be made.</returns>
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

        /// <summary>
        /// Builds the Datatable that shows which parts have a Restock Sales Order made and are just waiting to be built.  USes the Min-Max DataTable.
        /// </summary>
        /// <param name="pendingDt">Current DataTable where information will be stored.</param>
        /// <param name="minMaxDt">Min-Max DataTable.</param>
        /// <returns>A DataTable where the parts are waiting to be built.</returns>
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

        /// <summary>
        /// Helper method used to determine if the current date falls within 45 days of a quarter end.
        /// </summary>
        /// <returns>True if the date is within 45 days of an EOQ, False otherwise.</returns>
        private static bool IsDuringQuarterlyRush()
        {
            // Year does not matter, 2018 is only used because DateTime constructor requires a date.
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

        /// <summary>
        /// Helper Class for the IsDuringQuarterlyRush method.
        /// </summary>
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

        /// <summary>
        /// Writes a DataTable to a CSV file at the specified location.  This method is used for testing only.
        /// </summary>
        /// <param name="dataTable">DataTable to be written to a CSV file.</param>
        /// <param name="filepath">Location of CSV file.</param>
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
