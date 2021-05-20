using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolConvertSqlToMongoDB
{
    class DBConnect
    {
        public static SqlConnection GetDBConnection(string datasource, string database, string username, string password)
        {
            //
            // Data Source=TRAN-VMWARE\SQLEXPRESS;Initial Catalog=simplehr;Persist Security Info=True;User ID=sa;Password=12345
            //
            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;
            //connString = "Server=tcp:futecheazyerp.database.windows.net,1433;Initial Catalog=Futech_Eazyerp2014_2021;Persist Security Info=False;User ID=eazyerp;Password=Futecheazy1357;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
            SqlConnection conn = new SqlConnection(connString);
           
            return conn;
        }

        /// <summary>
        /// xóa linked
        /// </summary>
        /// <param name="datasource"></param>
        public static void RemoveLinked(string datasource, SqlConnection conn)
        {
            try
            {

                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }

                var query = new StringBuilder();
                query.AppendLine(string.Format("IF EXISTS(SELECT * FROM sys.servers WHERE name = N'{0}')", datasource));
                query.AppendLine(string.Format("EXEC master.sys.sp_dropserver '{0}','droplogins'", datasource));
                //  query.AppendLine(" GO");

                SqlCommand cmd = new SqlCommand(query.ToString(), conn);

                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                conn.Close();
            }

        }

        /// <summary>
        /// tạo liên kết đến server thứ 2 và kiểm tra kết nối
        /// </summary>
        /// <param name="datasource"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="conn"></param>
        /// <returns></returns>
        public static bool CheckConnection(string datasource, string username, string password, SqlConnection conn)
        {
            bool ischeck = true;
            try
            {
                conn.Open();
                var query = new StringBuilder();
                //Tạo linked tại server 1 đến server thứ 2
              // query.AppendLine(string.Format("if not exists(select * from master.sys.servers where name = N'{0}')", datasource));
               query.AppendLine("if not exists(select * from tcp:futecheazyerp.database.windows.net.sys.sp_dropserver)");
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


                SqlCommand cmd = new SqlCommand(query.ToString(), conn);

                cmd.ExecuteNonQuery();

            }
            catch (Exception ex )
            {
                ischeck = false;
                var s = ex.Message;
                //xóa linked
                DBConnect.RemoveLinked(datasource, conn);
                throw;
            }
            finally
            {
                conn.Close();
            }

            return ischeck;
        }

        public static DataSet GetDataSet(string strSQL, SqlConnection conn, bool isStoredProcedure = false)
        {
            //Khai báo 1 biến bảng để chứa dữ liệu
            var ds = new DataSet();

            try
            {
                //Mở kết nối
                conn.Open();
                //Khai báo 1 đối tượng để thực hiện công việc
                SqlCommand comm = new SqlCommand(strSQL, conn);
                //SqlCommand comm = new SqlCommand();

                //Nếu sử dụng thủ tục
                if (isStoredProcedure)
                {
                    comm.CommandType = CommandType.StoredProcedure;
                }
                else
                {
                    comm.CommandType = CommandType.Text;
                }
                //Khai báo 1 đối tượng để chứa dữ liệu tạm thời lấy được từ db lên
                SqlDataAdapter adapter = new SqlDataAdapter(comm);
                //Đổ dữ liệu vào bảng

                adapter.Fill(ds);

                adapter.Dispose();
                comm.Dispose();
                conn.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //Đóng kết nối
                conn.Close();
            }

            return ds;
        }
        //public static DataSet GetDataSet(string strSQL, SqlConnection conn, bool isStoredProcedure = false)
        //{
        //    //Khai báo 1 biến bảng để chứa dữ liệu
        //    var ds = new DataSet();

        //    try
        //    {
        //        //Mở kết nối
        //        conn.Open();
        //        //Khai báo 1 đối tượng để thực hiện công việc
        //        SqlCommand comm = new SqlCommand(strSQL, conn);
        //        //SqlCommand comm = new SqlCommand();

        //        //Nếu sử dụng thủ tục
        //        if (isStoredProcedure)
        //        {
        //            comm.CommandType = CommandType.StoredProcedure;
        //        }
        //        else
        //        {
        //            comm.CommandType = CommandType.Text;
        //        }
        //        //Khai báo 1 đối tượng để chứa dữ liệu tạm thời lấy được từ db lên
        //        SqlDataAdapter adapter = new SqlDataAdapter(comm);
        //        //Đổ dữ liệu vào bảng

        //        adapter.Fill(ds);

        //        adapter.Dispose();
        //        comm.Dispose();
        //        conn.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //    finally
        //    {
        //        //Đóng kết nối
        //        conn.Close();
        //    }

        //    return ds;
        //}
        public static DataTable GetTable(string strSQL, string connect, bool isStoredProcedure = false)
        {
            //Khai báo 1 biến bảng để chứa dữ liệu
            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(connect))
            {
                try
                {
                    //Mở kết nối
                    conn.Open();
                    //Khai báo 1 đối tượng để thực hiện công việc
                    //SqlCommand comm = new SqlCommand(strSQL, conn);
                    SqlCommand comm = new SqlCommand();
                    //Thực hiện công việc trên database nào
                    comm.Connection = conn;
                    //Nếu sử dụng thủ tục
                    if (isStoredProcedure)
                    {
                        comm.CommandType = CommandType.StoredProcedure;
                    }
                    else
                    {
                        comm.CommandType = CommandType.Text;
                    }
                    //Công việc cần thực hiện
                    comm.CommandText = strSQL;
                    //"select * from [125.212.226.137].[MPARKING-TH2].dbo.[tblCardGroup]\r\n"
                    //Khai báo 1 đối tượng để chứa dữ liệu tạm thời lấy được từ db lên
                    SqlDataAdapter adapter = new SqlDataAdapter(comm);
                    //Đổ dữ liệu vào bảng
                    adapter.Fill(dt);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    //Đóng kết nối
                    conn.Close();
                }
            }
            return dt;
        }
    }
}
