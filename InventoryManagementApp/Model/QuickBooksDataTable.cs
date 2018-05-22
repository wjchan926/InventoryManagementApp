﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;

namespace InventoryManagementApp.Model
{
    abstract class QuickBooksDataTable : DataTable, IQuickBooksData
    {
        protected abstract string sqlCmdStr { get; }
              
        public void BuildTable()
        {
            QueryQB(sqlCmdStr);
        }

        protected void QueryQB(string sqlCmdStr)
        {
   //         ConsoleWriter.WriteLine("Accessing QuickBooks Database.");
            
            using (OdbcConnection con = new OdbcConnection("Dsn=QuickBooks Data"))
            {
                con.Open(); //Open Connection
                
                using (OdbcDataAdapter dAdapter = new OdbcDataAdapter(sqlCmdStr, con))
                {
                    dAdapter.FillError += new FillErrorEventHandler(FillError);
                    try
                    {
                        dAdapter.Fill(this);
    //                    ConsoleWriter.WriteLine("Data Table Filled.");
                    }
                    catch (OdbcException sqlError)
                    {
  //                      ConsoleWriter.WriteLine("SQL Statment Incorrect: " + sqlError.Message);
                    }
                    catch (Exception)
                    {
    //                    ConsoleWriter.WriteLine("Data Table Filled Failed.");
                    }
                }
            }
             
        }

        protected abstract void FillError(object sender, FillErrorEventArgs args);     
    }
}
