using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;

namespace InventoryManagementApp
{
    sealed class ItemDataTable : QuickBooksDataTable
    {               
        public override string sqlCmdStr
        {
            get
            {
                return "SELECT ItemInventoryAssembly.FullName AS PartNumber, ItemInventoryAssembly.QuantityOnHand AS QtyOnHand " +
                             "FROM ItemInventoryAssembly";
            }
        }
                
        protected override sealed void FillError(object sender, FillErrorEventArgs args)
        {
            // Code to handle precision loss.  
            object errorarg = DBNull.Value;
                             
            try
            {
                errorarg = Convert.ToDecimal(args.Values[1]);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Cannot convert value to Decimal.\n" + e.Message);
            }

            DataRow myRow = args.DataTable.Rows.Add(new object[]
                {args.Values[0], errorarg});

            args.Continue = true;
        }
    }
}
