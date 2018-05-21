using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using System.Collections;
using System.Threading;
using System.Data;

namespace InventoryManagementApp
{
    [TestFixture]
    public class InvManTest
    {
        private ExcelDoc excelDoc;

        public InvManTest()
        {          
            
        }        
        
        [Test]
        public void InStreamDataTest()
        {
            excelDoc = new ExcelDoc();

            excelDoc.Open();
            excelDoc.InStreamData();
            excelDoc.Close();

            StringBuilder sb = new StringBuilder();

            foreach(KeyValuePair<string,int> kvp in excelDoc.partNumList)
            {
                sb.AppendLine(kvp.Key + " " + kvp.Value);
            }

            System.IO.File.WriteAllText(@"\\msw-fp1\user$\wchan\Documents\Visual Studio 2015\Projects\InventoryManagement\InventoryManagement\bin\Debug\Test\Part Numbers.txt", sb.ToString());

        } 
        
        [Test]
        public void PolyItemTest()
        {
            QuickBooksDataTable itemTable = new ItemDataTable();
            itemTable.BuildTable();

            itemTable.Write(@"\\msw-fp1\user$\wchan\Documents\Visual Studio 2015\Projects\InventoryManagement\InventoryManagement\bin\Debug\Test\PolyItem.csv");
        }

        [Test]
        public void PolySOTest()
        {
            QuickBooksDataTable soTable = new SODataTable();
            soTable.BuildTable();

            soTable.Write(@"\\msw-fp1\user$\wchan\Documents\Visual Studio 2015\Projects\InventoryManagement\InventoryManagement\bin\Debug\Test\PolySO.csv");
        }

        [Test]
        public void PolyMinMaxTest()
        {
            QuickBooksDataTable itemTable = new ItemDataTable();
            QuickBooksDataTable soTable = new SODataTable();
            itemTable.BuildTable();
            soTable.BuildTable();

            using (excelDoc = new ExcelDoc())
            {
                excelDoc.Open();
                excelDoc.InStreamData();
                excelDoc.Close();

                DataTable minMaxDt = new DataTable().BuildTable(soTable, itemTable, excelDoc.partNumList);
              
                minMaxDt.Write(@"\\msw-fp1\user$\wchan\Documents\Visual Studio 2015\Projects\InventoryManagement\InventoryManagement\bin\Debug\Test\PolyMinMax.csv");
            }            
        }

        [Test]
        public void ExcelWriteTest()
        {
            QuickBooksDataTable itemTable = new ItemDataTable();
            QuickBooksDataTable soTable = new SODataTable();
            itemTable.BuildTable();
            soTable.BuildTable();

            using (excelDoc = new ExcelDoc())
            {
                excelDoc.Open();
                excelDoc.InStreamData();
   
                DataTable minMaxDt = new DataTable().BuildTable(soTable, itemTable, excelDoc.partNumList);

                minMaxDt.Write(@"\\msw-fp1\user$\wchan\Documents\Visual Studio 2015\Projects\InventoryManagement\InventoryManagement\bin\Debug\Test\PolyMinMax.csv");

                excelDoc.Write(minMaxDt);
            }
        }
    }
}
