using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InventoryManagementApp.Model;
using System.Data;

namespace InventoryManagementApp.ViewModel
{
    class ExcelDocViewModel : IExcelViewModel
    {
        private ExcelDoc excelDoc = new ExcelDoc();
        
        private DataTable minMaxDt;
                        
        public void Open()
        {            
            excelDoc.Open();            
        }

        public void Close()
        {
            excelDoc.Close(); 
        }

        public void Dispose()
        {
            using (excelDoc)
            {
                excelDoc.Dispose();
            }
        }
        
        public void UpdateSO(DataTable soReqDataTable)
        {
            if (!excelDoc.excelObjSet)
            {
                excelDoc.SetExcelObjects();
            }
            excelDoc.UpdateSO(soReqDataTable); 
        }

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
