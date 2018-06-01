using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;

namespace InventoryManagementApp.Model
{
    /// <summary>
    /// An abstract super class that represents a DataTable from QuickBooks.
    /// </summary>
    abstract class QuickBooksDataTable : DataTable, IQuickBooksData
    {
        protected abstract string sqlCmdStr { get; }
        
        /// <summary>
        /// Builds the DataTable from the QuickBooks Database.
        /// </summary>                      
        public void BuildTable()
        {
            QueryQB(sqlCmdStr);
        }

        /// <summary>
        /// Queries the data from the QuickBooks Database.
        /// </summary>
        /// <param name="sqlCmdStr">SQL String used for querying data</param>
        protected void QueryQB(string sqlCmdStr)
        {    
            try
            {
                using (OdbcConnection con = new OdbcConnection("Dsn=QuickBooks Data"))
                {
                    con.Open(); //Open Connection
                    Log.WriteLine("Accessing QuickBooks Database.");

                    using (OdbcDataAdapter dAdapter = new OdbcDataAdapter(sqlCmdStr, con))
                    {
                        dAdapter.FillError += new FillErrorEventHandler(FillError);
                        try
                        {
                            dAdapter.Fill(this);
                        }
                        catch (OdbcException sqlError)
                        {
                            Console.WriteLine("SQL Statment Incorrect: " + sqlError.Message);
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("Data Table Filled Failed.");
                        }
                    }
                }
            }
            catch 
            {
                Log.WriteLine("QuickBooks Connection Failed.");
            }
             
        }

        /// <summary>
        /// Error handler if the returned data cannot be converted by .NET from SQL.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        protected abstract void FillError(object sender, FillErrorEventArgs args);     
    }
}
