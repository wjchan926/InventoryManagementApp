using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections;
using System.Diagnostics;
using System.Data;

namespace InventoryManagementApp.Model
{
    class ExcelDoc : IDisposable
    {
        public Excel.Application myApp { get; private set; } 
        public Excel.Workbook myBook { get; private set; } 
        public Excel.Workbooks myBooks { get; private set; }
        public Excel.Worksheet mySheet { get; private set; }
        public Excel.Range myRange { get; private set; }
        public readonly string minMaxPath = @"\\MSW-FP1\Shared\DG Inventory Management.xlsx";
        public Dictionary<string, ExcelPartNumber> partNumList { get; private set; }
        public bool excelObjSet { get; private set; } = false;
        
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
                    Log.WriteLine("Instance of Excel Found");
                }
                catch (COMException e)
                {
                    Console.WriteLine("No Instance of Excel Found:\n" + e.Message);
                }
            }
            else
            {
                try
                {
                    myApp = new Excel.Application();
                    Log.WriteLine("New Instance of Excel Created.");
                }
                catch (Exception)
                {
                    Log.WriteLine("Cannot Access File on network, try Again.");
                }
            }

            myApp.Visible = true;            // True to see new instance, false to hide
            myApp.DisplayAlerts = false;            // Hide alerts

            // Set the objects to corresponding excel objects
            myBooks = myApp.Workbooks;
            myBook = myBooks.Open(minMaxPath);
            mySheet = myBook.Sheets["Marlin Steel"];
            
            // SetExcelObjects();
            Log.WriteLine("Min-Max Document Opened.");
        }
        
        /// <summary>
        /// Sets the excel objects
        /// </summary>
        public void SetExcelObjects()
        {
            // Sets workbook to path specified                  
            try
            {
                myApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                myBooks = myApp.Workbooks;
                myBook = myBooks["DG Inventory Management.xlsx"];
                mySheet = myBook.Sheets["Marlin Steel"];
                setRange();
                excelObjSet = true;
            }
            catch (NullReferenceException e)
            {
                // If file is not found
                Console.WriteLine(e.Message);
            }
            catch (Exception e)
            {
                // Other problemsW
                Console.WriteLine(e.Message);
            }
            Log.WriteLine("Excel Objects Set.");
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

            myRange = mySheet.Range[ExcelColumn.partNumber + "2", ExcelColumn.restockSODate + lastUsedRow + ""];        
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
                excelObjSet = false;
                Log.WriteLine("Min-Max Document Saved and Closed.");
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
            partNumList = new Dictionary<string, ExcelPartNumber>();

            foreach (Excel.Range row in myRange.Rows)
            {
                object value = myRange[row.Row-1, ExcelColumn.partNumber].Value2;
                string convertedPartNumber = Convert.ToString(value);
                //     partNumList.Add(convertedPartNumber, row.Row);
                dynamic soNumVal = myRange[row.Row-1, ExcelColumn.restockSONum].Value2;
                string conSoNumVal = Convert.ToString(soNumVal);
                dynamic soDateVal = myRange[row.Row-1, ExcelColumn.restockSODate].Value2;
                string conSoDateVal = Convert.ToString(soDateVal);
                partNumList.Add(convertedPartNumber, new ExcelPartNumber(row.Row-1, conSoNumVal, conSoDateVal));
            }
            Log.WriteLine(partNumList.Count + " Entries Found.");       
        }

        public void Write(object writeOb)
        {
            DataTable minMaxDt = (DataTable)writeOb;
            
            foreach (DataRow row in minMaxDt.Rows)
            {                
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.min] = row["Min"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.max] = row["Max"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.onHand] = row["QtyOnHand"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.avgSalePrice] = String.Format("{0:C}", row["AvgSalePrice"]);
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.quantitySold] = row["Last15Months"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.maxStockRev] = String.Format("{0:C}", row["MaxStockRev"]);
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.restockSONum] = row["RestockSONum"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.restockSODate] = row["RestockSODate"];
                Log.WriteLine(row["PartNumber"].ToString() + " Analyzed");
            }

            Log.WriteLine("Analysis Complete.");
        }   
        
        public void UpdateSO(DataTable soReqDataTable)
        {
            foreach (DataRow row in soReqDataTable.Rows)
            {
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.restockSONum] = row["RestockSONum"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.restockSODate] = row["RestockSODate"];             
            }
            Log.WriteLine("Restock SO Updated on Min-Max Document.");
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
                Log.WriteLine("All Excel Objects Released.");
                excelObjSet = false;
            }
            catch (Exception e)
            {
                Console.WriteLine("Release Failed:\n" + e.Message);
            }
        }

    }

    class ExcelPartNumber
    {
        public int rowNum;
        public string restockSONum;
        public string restockSODate;

        public ExcelPartNumber(int rowNum, string restockSONum, string restockSODate)
        {
            this.rowNum = rowNum;
            this.restockSONum = restockSONum;
            this.restockSODate = restockSODate;
        }
    }

}
