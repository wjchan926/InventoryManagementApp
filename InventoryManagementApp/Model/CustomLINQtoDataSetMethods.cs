using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace InventoryManagementApp.Model
{
    /// <summary>
    /// Used for the CopyDataTable() for SQL statements that include things like LEFT OUTER JOIN
    /// </summary>
    public static class CustomLINQtoDataSetMethods
    {
        public static DataTable CustomCopyToDataTable<T>(this IEnumerable<T> source)
        {
            return new ObjectShredder<T>().Shred(source, null, null);
        }

        public static DataTable CustomCopyToDataTable<T>(this IEnumerable<T> source,
                                                    DataTable table, LoadOption? options)
        {
            return new ObjectShredder<T>().Shred(source, table, options);
        }

    }
}
