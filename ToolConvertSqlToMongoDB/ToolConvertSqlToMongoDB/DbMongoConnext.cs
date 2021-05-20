using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolConvertSqlToMongoDB
{
    class DbMongoConnext
    {
        /// <summary>
        /// tạo liên kết đến server thứ 2 và kiểm tra kết nối
        /// </summary>
        /// <param name="datasource"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="conn"></param>
        /// <returns></returns>
        public static bool CheckConnection(string datasource, string username, string password)
        {
            bool ischeck = true;
            try
            {
                //conn.Open();
                var query = new StringBuilder();
                //Tạo linked tại server 1 đến server thứ 2
                query.AppendLine(string.Format("if not exists(select * from sys.servers where name = N'{0}')", datasource));
                query.AppendLine("BEGIN");
                query.AppendLine("EXEC sp_addlinkedserver");
                query.AppendLine(string.Format("@server='{0}',", datasource));
                query.AppendLine("@srvproduct='',");
                query.AppendLine("@provider='sqlncli',");
                query.AppendLine(string.Format("@datasrc='{0}',", datasource));
                query.AppendLine("@location='',");
                query.AppendLine("@provstr=''");
                query.AppendLine("EXEC sp_addlinkedsrvlogin");
                query.AppendLine(string.Format("@rmtsrvname = '{0}',", datasource));
                query.AppendLine("@useself = 'false',");
                query.AppendLine(string.Format("@rmtuser = '{0}',", username));
                query.AppendLine(string.Format("@rmtpassword = '{0}'", password));
                query.AppendLine("END");

                //test kết nối
                query.AppendLine("declare @srvr nvarchar(128), @retval int;");
                query.AppendLine(string.Format("set @srvr = '{0}';", datasource));
                query.AppendLine("begin try");
                query.AppendLine("    exec @retval = sys.sp_testlinkedserver @srvr;");
                query.AppendLine("end try");
                query.AppendLine("begin catch");
                query.AppendLine("  set @retval = sign(@@error);");
                query.AppendLine("end catch;");
                query.AppendLine("if @retval <> 0");
                query.AppendLine("raiserror('Kết nối không thành công. Vui lòng kiểm tra lại thông tin kết nối!', 16, 2 );");



            }
            catch (Exception)
            {
                ischeck = false;
                //xóa linked
                //DBConnect.RemoveLinked(datasource, conn);
                throw;
            }
            //finally
            //{
            //    conn.Close();
            //}

            return ischeck;
        }
       
    }
}
