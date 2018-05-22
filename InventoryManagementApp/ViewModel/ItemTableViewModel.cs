using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InventoryManagementApp.Model;
using System.Data;

namespace InventoryManagementApp.ViewModel
{
    class ItemTableViewModel
    {
        public IQuickBooksData itemDataTable
        {
            get
            {
                return new ItemDataTable();
            }        
        }
    }
}
