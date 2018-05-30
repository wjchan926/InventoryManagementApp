using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;
using System.Collections;

namespace InventoryManagementApp.Model
{
    sealed class SODataTable : QuickBooksDataTable, IQuickBooksData
    {       
        protected override string sqlCmdStr
        {
            get
            {
                return "SELECT SalesOrderLine.SalesOrderLineItemRefFullName AS PartNumber, SalesOrderLine.SalesOrderLineQuantity AS Quantity, SalesOrderLine.SalesOrderLineRate AS SalePrice, SalesOrderLine.ShipDate AS ShipDate " +
                           "FROM SalesOrderLine " +
                           "WHERE (SalesOrderLine.SalesOrderLineItemRefFullName IS NOT NULL) AND (SalesOrderLine.ShipDate IS NOT NULL) AND (SalesOrderLine.SalesOrderLineQuantity IS NOT NULL) AND (SalesOrderLine.SalesOrderLineRate > 0)";
            }
        }
        
        protected override sealed void FillError(object sender, FillErrorEventArgs args)
        {
            // Code to handle precision loss.  
            object errorSOQty = 0.0m;
            object errorSalePrice = 0.0m;
            object errorShipDate = DBNull.Value;

            try
            {
                errorSOQty = Convert.ToDecimal(args.Values[1]);
                errorSalePrice = Convert.ToDecimal(args.Values[2]);
                errorShipDate = DateTime.Parse(args.Values[3].ToString());
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Cannot convert SO Data.\n" + e.Message);
                System.Diagnostics.Debug.WriteLine(errorSOQty.ToString() + " " + errorSalePrice.ToString() + " " + errorShipDate.ToString());
            }

            DataRow myRow = args.DataTable.Rows.Add(new object[]
                {args.Values[0], errorSOQty, errorSalePrice, errorShipDate});

            args.Continue = true;
        }

    }
}
