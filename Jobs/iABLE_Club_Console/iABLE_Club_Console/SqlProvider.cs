using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iABLE_Club_Console
{
    public class SqlProvider
    {
        public static DataTable dbGetDataTable(string connectionName, string sql)
        {
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(sql, conn);
            da.SelectCommand.CommandTimeout = 300;
            try
            {
                da.Fill(dt);
            }
            catch
            {
                da.Dispose();
                conn.Close();
                conn.Dispose();
            }
            return dt;
        }

        public static int dbExecuteNonQuery(string connectionName, string sql)
        {
            int i = 0;
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                conn.Open();
                i = cmd.ExecuteNonQuery();
            }
            catch
            {

            }
            finally
            {
                cmd.Cancel();
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return i;
        }

        public static object dbExecuteScalar(string connectionName, string sql)
        {
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);

            try
            {
                conn.Open();
                return cmd.ExecuteScalar();
            }
            catch
            {
                return null;
            }
            finally
            {
                cmd.Cancel();
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
    }
}
