using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iABLE_Club_Console
{
    public class OraDBUtil
    {
        public static DataTable dbGetDataTable(string connectionName, string sqlCmd)
        {

            OracleConnection g_adoConn = new OracleConnection(ConfigurationManager.ConnectionStrings[connectionName].ConnectionString);
            DataTable dt = new DataTable();
            OracleDataAdapter da = new OracleDataAdapter(sqlCmd, g_adoConn);
            da.SelectCommand.CommandTimeout = 300;
            try
            {
                da.Fill(dt);
            }
            catch(Exception ex)
            {
                da.Dispose();
                g_adoConn.Close();
                g_adoConn.Dispose();
                throw ex;
            }
            g_adoConn.Close();
            return dt;
        }

    //    Dim g_adoConn As New OracleConnection(ConfigurationManager.ConnectionStrings(ConnectionName).ConnectionString)
    //    Dim dt As New DataTable
    //    Dim da As New OracleDataAdapter(strSqlCmd, g_adoConn)

    //    If CommandTimeoutInSecond > 0 Then
    //        da.SelectCommand.CommandTimeout = CommandTimeoutInSecond
    //    End If
    //    Threading.Thread.CurrentThread.Priority = Threading.ThreadPriority.BelowNormal
    //    Try
    //        da.Fill(dt)
    //    Catch ex As Exception
    //        g_adoConn.Close() : g_adoConn.Dispose()
    //        Threading.Thread.CurrentThread.Priority = Threading.ThreadPriority.Normal
    //        Throw ex
    //    End Try
    //    g_adoConn.Close() : g_adoConn = Nothing
    //    Threading.Thread.CurrentThread.Priority = Threading.ThreadPriority.Normal
    //    Return dt
    //End Function
    }
}
