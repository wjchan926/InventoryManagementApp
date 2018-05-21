﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;
using System.Diagnostics;
using System.Data;

namespace InventoryManagementApp
{
    class ExcelDoc : IDisposable
    {
        public Excel.Application myApp { get; private set; } 
        public Excel.Workbook myBook { get; private set; } 
        public Excel.Workbooks myBooks { get; private set; }
        public Excel.Worksheet mySheet { get; private set; }
        public Excel.Range myRange { get; private set; }
        public readonly string minMaxPath = @"\\msw-fp1\shared\DG Inventory Management.xlsx";
        public Dictionary<string, int> partNumList { get; private set; }
        
        /// <summary>
        /// Default Constructor for the Excel Doc class
        /// </summary>
        public ExcelDoc()
        {
            myApp = null;
            myBooks = null;
            myBook = null;
            mySheet = null;
            myRange = null;
            partNumList = new Dictionary<string, int>();
        }

        /// <summary>
        /// Opens the MinMax Document
        /// </summary>
        public void Open()
        {
            if (Process.GetProcessesByName("EXCEL").Count() > 0)
            {
                // Creates a new instance of excel
                try
                {
                    myApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    Console.WriteLine("Instance of Excel Found");
                }
                catch (COMException e)
                {
                    Console.WriteLine("No Instance of Excel Found:\n" + e.Message);
                }
            }
            else
            {
                myApp = new Excel.Application();
                Console.WriteLine("New Instance of Excel Created.");
            }

            myApp.Visible = true;            // True to see new instance, false to hide
            myApp.DisplayAlerts = false;            // Hide alerts

            // Set the objects to corresponding excel objects
            SetExcelObjects(); 
        }
        
        /// <summary>
        /// Sets the excel objects
        /// </summary>
        public void SetExcelObjects()
        {
            // Sets workbook to path specified                  
            try
            {
                myBooks = myApp.Workbooks;
                myBook = myBooks.Open(minMaxPath);
                mySheet = myBook.Sheets["Marlin Steel"];
                setRange();
            }
            catch (NullReferenceException e)
            {
                // If file is not found
                Console.WriteLine(e.Message);
                throw;
            }
            catch (Exception e)
            {
                // Other problems
                Console.WriteLine(e.Message);
                throw;
            }
        }

        /// <summary>
        /// Sets the range of of jobs to analyze with current selection
        /// </summary>
        private void setRange()
        {
            // All Cells used in column A and set as range     
            Excel.Range lastCell = mySheet.Cells;
            Excel.Range lastCellUsed = lastCell.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int lastUsedRow = lastCellUsed.Row;

            Marshal.ReleaseComObject(lastCellUsed);
            Marshal.ReleaseComObject(lastCell);

            myRange = mySheet.Range["A2", ExcelColumn.maxStockRev + lastUsedRow];
        }
        
        /// <summary>
        /// Closes workbook and applicaiton.  Releases Objects
        /// Called with the Save and Close GUI Method
        /// </summary>
        public void Close()
        {
            try
            {
                myBook.Close(true, Type.Missing, Type.Missing);
                myBooks.Close();
                myApp.Quit();
                myApp.DisplayAlerts = true;
                Console.WriteLine("Excel Closed.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);    
            }
        }
                         
        /// <summary>
        /// Gets all the part numbers from the excel spreasheet
        /// </summary>
        public void InStreamData()
        { 
            foreach (Excel.Range row in myRange.Rows)
            {
                dynamic value = row.Value2;
                string convertedPartNumber = Convert.ToString(value);
                partNumList.Add(convertedPartNumber, row.Row);
            }
            Console.WriteLine(partNumList.Count + " Entries Found.");       
        }

        public void Write(object writeOb)
        {
            DataTable minMaxDt = (DataTable)writeOb;

            foreach (KeyValuePair<string, int> kvp in partNumList)
            {
                
                mySheet.Cells[kvp.Value, ExcelColumn.min] = minMaxDt.Rows.Find(kvp.Key)["Min"];
                mySheet.Cells[kvp.Value, ExcelColumn.max] = minMaxDt.Rows.Find(kvp.Key)["Max"];
                mySheet.Cells[kvp.Value, ExcelColumn.onHand] = minMaxDt.Rows.Find(kvp.Key)["QtyOnHand"];
                mySheet.Cells[kvp.Value, ExcelColumn.avgSalePrice] = minMaxDt.Rows.Find(kvp.Key)["AvgSalePrice"];
                mySheet.Cells[kvp.Value, ExcelColumn.quantitySold] = minMaxDt.Rows.Find(kvp.Key)["Last15Months"];
                mySheet.Cells[kvp.Value, ExcelColumn.maxStockRev] = minMaxDt.Rows.Find(kvp.Key)["MaxStockRev"];
                //                ConsoleWriter.WriteLine(kvp.Key + " Data Written.");           
            }
        }   
        
        public void Dispose()
        {
            try
            {
                Marshal.ReleaseComObject(myRange);
                Marshal.ReleaseComObject(mySheet);
                Marshal.ReleaseComObject(myBook);
                Marshal.ReleaseComObject(myBooks);
                Marshal.ReleaseComObject(myApp);
                Console.WriteLine("All Excel Objects Released.");
            }
            catch (Exception e)
            {
                Console.WriteLine("Release Failed:\n" + e.Message);
            }
        }
    }
}
