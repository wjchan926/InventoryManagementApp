using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InventoryManagementApp.Model;
using System.Data;

namespace InventoryManagementApp.ViewModel
{
    /// <summary>
    /// ViewModel that binds the Excel Min-Max Document to their respective DataTabgles
    /// </summary>
    class ExcelDocViewModel : IExcelViewModel
    {
        private ExcelDoc excelDoc = new ExcelDoc();
        
        private DataTable minMaxDt;
         
        /// <summary>
        /// Opens the Excel Document
        /// </summary>                               
        public void Open()
        {            
            excelDoc.Open();            
        }

        /// <summary>
        /// Closes the Excel Document
        /// </summary>
        public void Close()
        {
            excelDoc.Close(); 
        }

        /// <summary>
        /// Releases all Excel Document COM Objects
        /// </summary>
        public void Dispose()
        {
            using (excelDoc)
            {
                excelDoc.Dispose();
            }
        }

        /// <summary>
        /// Updates that Sales Order Required Dates in the Excel Min-Max Document that is data-bound to the soReqDataTable.
        /// </summary>
        /// <param name="soReqDataTable"></param>
        public void UpdateSO(DataTable soReqDataTable)
        {
            if (!excelDoc.excelObjSet)
            {
                excelDoc.SetExcelObjects();
            }

            excelDoc.UpdateSO(soReqDataTable); 
        }

        /// <summary>
        /// Builds the Min-Max DataTable.
        /// </summary>
        /// <returns>A DataTable representing all the QuickBooks Data to be written to the Min-Max Doc.</returns>
        public DataTable Analyze()
        {         
            if (!excelDoc.excelObjSet)
            {
                excelDoc.SetExcelObjects();
            }

            if (excelDoc.excelObjSet)
            {
                Log.WriteLine("...Analyzing Part Numbers...");

                IQuickBooksData itemDataTable = new ItemDataTable();
                IQuickBooksData salesOrderDataTable = new SODataTable();

                excelDoc.InStreamData();
                itemDataTable.BuildTable();
                salesOrderDataTable.BuildTable();

                minMaxDt = new DataTable().BuildTable(salesOrderDataTable, itemDataTable, excelDoc.partNumList);

                excelDoc.Write(minMaxDt);
                return minMaxDt;
            }
            else
            {
                Log.WriteLine("Cannot Access Min-Max Document.");
                return new DataTable();
            }
      
        }    
    }
}
