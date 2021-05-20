using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolConvertSqlToMongoDB.Models
{
   public class SqlExQuery<T> where T : class
    {
        //public static List<T> ExcuteQuery(string storeName)
        //{
        //    using (var _DataContext = new DatabaseFactory().Get())
        //    {
        //        var result = _DataContext.Database.SqlQuery<T>(storeName);
        //        return result.ToList();
        //    }
        //}
    }
}
