using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using InventoryManagementApp.Model;

namespace InventoryManagementApp.ViewModel
{
    interface IExcelViewModel
    {        
        void Open();
        void Close();
        void Dispose();
        void UpdateSO(DataTable soReqDataTable);
        DataTable Analyze();
    }
}
