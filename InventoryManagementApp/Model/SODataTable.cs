using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;
using System.Collections;
using System.Windows;

namespace InventoryManagementApp.Model
{
    /// <summary>
    /// Represents the SalesOrderLine DataTable from QuickBooks
    /// </summary>
    sealed class SODataTable : QuickBooksDataTable, IQuickBooksData
    {
        /// <summary>
        /// SQL string property.
        /// </summary>
        protected override string sqlCmdStr
        {
            get
            {
                return "SELECT SalesOrderLine.SalesOrderLineItemRefFullName AS PartNumber, SalesOrderLine.SalesOrderLineQuantity AS Quantity, SalesOrderLine.SalesOrderLineRate AS SalePrice, SalesOrderLine.ShipDate AS ShipDate, SalesOrderLine.CustomerRefFullName AS Customer " +
                           "FROM SalesOrderLine " +
                           "WHERE (SalesOrderLine.SalesOrderLineItemRefFullName IS NOT NULL) AND (SalesOrderLine.ShipDate IS NOT NULL) AND (SalesOrderLine.SalesOrderLineQuantity IS NOT NULL) AND (SalesOrderLine.SalesOrderLineRate > 0)";
            }
        }

        /// <summary>
        /// Error handler if the returned data cannot be converted by .NET from SQL.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        protected override sealed void FillError(object sender, FillErrorEventArgs args)
        {
            // Code to handle precision loss.  
            object errorSOQty = 0.0m;
            object errorSalePrice = 0.0m;
            object errorShipDate = DBNull.Value;
        //    object errorShipDate = Convert.ToString((DateTime?)null);

            try
            {
                errorSOQty = Convert.ToDecimal(args.Values[1]);
                errorSalePrice = Convert.ToDecimal(args.Values[2]);
                errorShipDate = DateTime.Parse(args.Values[3].ToString());
              //  errorShipDate = DateTime.ParseExact(args.Values[3].ToString(),"MM-dd-yyyy",null);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Cannot convert SO Data.\n" + e.Message);
             //   MessageBox.Show(errorSOQty.ToString() + " " + errorSalePrice.ToString() + " " + errorShipDate.ToString());
            }

            DataRow myRow = args.DataTable.Rows.Add(new object[]
                {args.Values[0], errorSOQty, errorSalePrice, errorShipDate});

            args.Continue = true;
        }

    }
}
