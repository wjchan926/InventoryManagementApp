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
using System.IO;

namespace InventoryManagementApp.Model
{
    /// <summary>
    /// Represents all of the Excel Min-Max Document COM Objects.  Implements the IDisposable interface.
    /// </summary>
    public class ExcelDoc : IDisposable
    {
        Excel.Application myApp;
        Excel.Workbook myBook;
        Excel.Workbooks myBooks;
        Excel.Worksheet mySheet;
        Excel.Range myRange;
        readonly string minMaxPath = @"\\MSW-FP1\Shared\DG Inventory Management.xlsx";
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
        /// Opens the MinMax Document.
        /// </summary>
        public void Open()
        {
            try
            {
                if (Process.GetProcessesByName("EXCEL").Count() > 0)
                {
                    // Try to get an open instance of Excel.
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
                        // Try to Create new instance of Excel.
                        myApp = new Excel.Application();
                        Log.WriteLine("New Instance of Excel Created.");
                    }
                    catch (Exception)
                    {
                        Log.WriteLine("Cannot Access File on network, try Again.");
                    }
                }

                myApp.Visible = true;            // True to see new instance, false to hide


                // Set the objects to corresponding excel objects
                myBooks = myApp.Workbooks;
                myBook = myBooks.Open(minMaxPath);
                mySheet = myBook.Sheets["Marlin Steel"];
                setRange();

                // All excel objects are referenced
                excelObjSet = true;
                
                Log.WriteLine("Min-Max Document Opened.");
            }
            catch
            {
                Log.WriteLine("Cannot Access Min-Max Document.");
            }
        }
        
        /// <summary>
        /// Sets the excel objects to their corresponding COM objects.
        /// </summary>
        public void SetExcelObjects()
        {
            // Sets workbook to path specified    .              
            try
            {
                myApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                myBooks = myApp.Workbooks;
                myBook = myBooks["DG Inventory Management.xlsx"];
                mySheet = myBook.Sheets["Marlin Steel"];
                setRange();
                excelObjSet = true;
                Console.WriteLine("Excel Objects Set.");
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

        }

        /// <summary>
        /// Sets the range of of jobs to the entire list of jobs.
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
        /// Closes workbook and applicaiton.  Releases Objects.
        /// Called with the Save and Close GUI Method.
        /// </summary>
        public void Close()
        {
            try
            {
                myApp.DisplayAlerts = false;            // Hide alerts
                SetExcelObjects();
                myBook.Close(true, Type.Missing, Type.Missing);
                myBooks.Close();
                myApp.Quit();
                myApp.DisplayAlerts = true;
                Dispose();
                Log.WriteLine("Min-Max Document Saved and Closed.");    
            }
            catch (Exception e)
            {
                Log.WriteLine("Cannot Access Min-Max Document.");
                Console.WriteLine(e.Message);    
            }
        }
                         
        /// <summary>
        /// Gets all the part numbers from the excel spreasheet and stores them in partNumList.
        /// </summary>
        public void InStreamData()
        {
            try
            {
                partNumList = new Dictionary<string, ExcelPartNumber>();

                foreach (Excel.Range row in myRange.Rows)
                {
                    object value = myRange[row.Row - 1, ExcelColumn.partNumber].Value2;
                    string convertedPartNumber = Convert.ToString(value);
                    //     partNumList.Add(convertedPartNumber, row.Row);
                    dynamic soDateVal = myRange[row.Row - 1, ExcelColumn.restockSODate].Value2;
                    string conSoDateVal = Convert.ToString(soDateVal);
                    object bracketVal = myRange[row.Row - 1, ExcelColumn.bracketsPerSheet].Value2;
                    int convertedBracketPerSheet = Convert.ToInt32(bracketVal);
                    partNumList.Add(convertedPartNumber, new ExcelPartNumber(row.Row - 1, conSoDateVal, convertedBracketPerSheet));
                }
                Log.WriteLine(partNumList.Count + " Entries Found.");
            }
            catch 
            {
                Log.WriteLine("Cannot Access Min-Max Document.");
            }
        }

        /// <summary>
        /// Writes the Min-Max DataTable to the Excel Min-Max Document.
        /// </summary>
        /// <param name="writeOb">DatTable to write to the Excel Min-Max Document</param>
        public void Write(DataTable minMaxDt)
        {
            myApp.DisplayAlerts = false;            // Hide alerts
                
            foreach (DataRow row in minMaxDt.Rows)
            {
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.partNumber].Formula = HyperlinkPartNumber(row["PartNumber"].ToString());
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.min] = row["Min"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.max] = row["Max"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.onHand] = row["QtyOnHand"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.avgSalePrice] = String.Format("{0:C}", row["AvgSalePrice"]);
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.quantitySold] = row["Last15Months"];
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.maxStockRev] = String.Format("{0:C}", row["MaxStockRev"]);
                myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.restockSODate] = row["RestockSODate"];
                Log.WriteLine(row["PartNumber"].ToString() + " Analyzed");
            }

            Log.WriteLine("Analysis Complete.");

            myApp.DisplayAlerts = true;
        }   

        /// <summary>
        /// Hyperlinks the part number to the Release Document.  May not work if Release is named slightly different.
        /// </summary>
        /// <param name="partNumber">Part Number that corresponds to a Release Document</param>
        /// <returns>A string representing a forumla hyperlinking the part number to it's corresponding Release Document.  
        /// If it cannot be found, a string of the part nubmer is returned instead.</returns>
        private string HyperlinkPartNumber(string partNumber)
        {
            StringBuilder hyperlinkPartNumber = new StringBuilder();

            // Format the part number first
            // Cut off prefix
            // Cut off suffix
            int firstHyphen = partNumber.IndexOf('-');
            int lastHyphen = partNumber.LastIndexOf('-');

            StringBuilder formattedPartNumber = new StringBuilder((firstHyphen == lastHyphen ? partNumber.Substring(0, lastHyphen) : partNumber.Substring(firstHyphen+1, lastHyphen-firstHyphen-1)));
            
            if (partNumber.Length <= 10)
            {
                formattedPartNumber.Insert(0, "M");
            }

            return File.Exists(@"\\MSW-FP1\Factory\RELEASED DESIGNS\" + formattedPartNumber + ".pdf") ? 
                hyperlinkPartNumber.Append(@"=HYPERLINK(""\\MSW-FP1\Factory\RELEASED DESIGNS\" + formattedPartNumber + @".pdf"",""" + partNumber + @""")").ToString() :
                    hyperlinkPartNumber.Append(partNumber).ToString();

        }

        /// <summary>
        /// Writes the Sales Order Restock Date to Excel from the soReqDataTable View Model.
        /// </summary>
        /// <param name="soReqDataTable">DataTable that represents which part numbers need Restock Sales Orders</param>
        public void UpdateSO(DataTable soReqDataTable)
        {
            try
            {
                foreach (DataRow row in soReqDataTable.Rows)
                {
                    myRange[partNumList[row["PartNumber"].ToString()].rowNum, ExcelColumn.restockSODate] = String.Format("{0:M/d/yyyy}", row["RestockSODate"]);
                }
                Log.WriteLine("Restock SO Updated on Min-Max Document.");
            }
            catch
            {
                Log.WriteLine("Cannot Access Min-Max Document.");
            }
  
        }


        /// <summary>
        /// Releases all COM objects used in by the Excel document.  Required for the IDisposable interface.
        /// </summary>
        public void Dispose()
        {
            try
            {
                myApp.DisplayAlerts = true;
                Marshal.ReleaseComObject(myRange);
                Marshal.ReleaseComObject(mySheet);
                Marshal.ReleaseComObject(myBook);
                Marshal.ReleaseComObject(myBooks);
                Marshal.ReleaseComObject(myApp);
                Console.WriteLine("All Excel Objects Released.");
                excelObjSet = false;
                myApp = null;
                myBooks = null;
                myBook = null;
                mySheet = null;
                myRange = null;
            }
            catch (Exception e)
            {
                Log.WriteLine("Cannot Access Min-Max Document.");
                Console.WriteLine("Release Failed:\n" + e.Message);
            }
        }

    }

    /// <summary>
    /// Helper class to group data stored in partNumList.
    /// </summary>
   public class ExcelPartNumber
    {
        public int rowNum;
        public string restockSODate;
        public int bracketsPerSheet;

        public ExcelPartNumber(int rowNum, string restockSODate, int bracketsPerSheet)
        {
            this.rowNum = rowNum;
            this.restockSODate = restockSODate;
            this.bracketsPerSheet = bracketsPerSheet;
        }
    }

}
