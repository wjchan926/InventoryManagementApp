using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryManagementApp.Model
{
    /// <summary>
    /// Class that Logs status changes during application runtime.
    /// </summary>
    public static class Log
    {
        public static StringBuilder logSB = new StringBuilder();

        /// <summary>
        /// Writes a new line to the log.
        /// </summary>
        /// <param name="value">String to write to log</param>
        public static void WriteLine(string value)
        {
            logSB.AppendLine(value);
        }

        /// <summary>
        /// Clears all text from the log.
        /// </summary>
        public static void Clear()
        {
            logSB.Clear();
        }
    }
}
