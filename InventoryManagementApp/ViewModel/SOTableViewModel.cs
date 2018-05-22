using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InventoryManagementApp.Model;

namespace InventoryManagementApp.ViewModel
{
    class SOTableViewModel
    {
        public IQuickBooksData soDataTable
        {
            get
            {
                return new SODataTable();
            }
        }

    }
}
