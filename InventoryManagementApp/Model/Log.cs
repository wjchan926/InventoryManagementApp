using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagementApp.Model
{
    public static class Log
    {
        public static StringBuilder logSB = new StringBuilder();

        public static void WriteLine(string value)
        {
            logSB.AppendLine(value);
        }

        public static void Clear()
        {
            logSB.Clear();
        }
    }
}
