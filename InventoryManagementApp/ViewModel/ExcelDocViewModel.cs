﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InventoryManagementApp.Model;
using System.Data;

namespace InventoryManagementApp.ViewModel
{
    class ExcelDocViewModel
    {
        public ExcelDoc excelDoc { get; private set; }

        public ExcelDocViewModel()
        {
            excelDoc = new ExcelDoc();
        }
                
        public void Open()
        {
            excelDoc.Open();
        }

        public void Close()
        {
            excelDoc.Close();
        }

        public void Analyze(IQuickBooksData itemDataTable, IQuickBooksData salesOrderDataTable)
        {
            excelDoc.InStreamData();
            itemDataTable.BuildTable();
            salesOrderDataTable.BuildTable();

            DataTable minMaxDt = new DataTable().BuildTable(salesOrderDataTable, itemDataTable, excelDoc.partNumList);

            excelDoc.Write(minMaxDt);
        }

    }
}
