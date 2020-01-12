using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using InventoryManagementApp.Model;
using System.Text;
using System.Collections.Generic;

namespace InvManAppTest
{
    [TestClass]
    public class UnitTest1
    {

        private ExcelDoc excelDoc;

        [TestMethod]
        public void InStreamDataTest()
        {
            excelDoc = new ExcelDoc();

            excelDoc.Open();
            excelDoc.InStreamData();
            excelDoc.Close();

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, ExcelPartNumber> kvp in excelDoc.partNumList)
            {
                sb.AppendLine(kvp.Key + " " + kvp.Value.restockSODate + " " + kvp.Value.bracketsPerSheet);
            }

            System.IO.File.WriteAllText(@"\\msw-fp1\user$\wchan\Documents\InventoryManagementAppTest\Part Numbers.txt", sb.ToString());

        }
    }
}
