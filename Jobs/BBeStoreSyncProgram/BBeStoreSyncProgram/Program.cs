using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace BBeStoreSyncProgram
{
    class Program
    {
        static void Main(string[] args)
        {
            BBorderAPI.BBorderAPI api = new BBorderAPI.BBorderAPI();
            api.Timeout = 100000;
            try
            {
                api.AutoMaticallyTransfereStoreOrderToSAP();
            }
            catch (Exception ex)
            {
                string _smtp = ConfigurationManager.AppSettings.Get("MasterSMTP");
                string _from = ConfigurationManager.AppSettings.Get("MailFrom");
                string _to = ConfigurationManager.AppSettings.Get("MailTo");
                SmtpClient smtpClient1 = new SmtpClient(_smtp); 
                MailMessage mail = new MailMessage(_from, _to);
                mail.Subject = "Auto sync BB eStore failed";
                mail.IsBodyHtml = true;
                mail.Body = ex.ToString();
                try
                {
                    smtpClient1.Send(mail);
                }
                catch
                {

                }
            }


            try
            {
                DataTable dt = GetDataTable(CenterLibrary.DBConnection.eStoreBB, @"declare @start as datetime
                    declare @end as datetime
                    set @end = DATEADD(hour, 2, GETDATE())
                    set @start = DATEADD(day, -1, @end)
                    select OrderNo, UserID from [Order] where storeid='ABB' and 
                    OrderStatus in ('Confirmed', 'Closed_Converted', 'ConfirmdButNeedTaxIDReview', 'ConfirmdButNeedFreightReview')
                    and Orderdate between @start and @end");

                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string orderno = dr[0].ToString().Trim().ToUpper();
                        string userid = dr[1].ToString().Trim();
                        object count = ExecuteScalar(CenterLibrary.DBConnection.MyAdvantechGlobal, string.Format("select count(*) from BB_ESTORE_ORDER where ORDER_NO = '{0}'", orderno));
                        if (count != null && int.Parse(count.ToString()) == 0)
                        {
                            object ERPID = ExecuteScalar(CenterLibrary.DBConnection.MyAdvantechGlobal, string.Format("SELECT TOP 1 sd.COMPANY_ID FROM SIEBEL_CONTACT sc INNER JOIN SAP_DIMCOMPANY sd ON sc.ERPID = sd.COMPANY_ID WHERE sc.EMAIL_ADDRESS = '{0}' AND sc.OrgID = 'ABB' AND sd.ORG_ID = 'US10' AND sd.COMPANY_TYPE ='Z001' AND sc.ACCOUNT_ROW_ID<>'1-1CK7KL5' ", userid));
                            if (ERPID != null && !string.IsNullOrEmpty(ERPID.ToString()))
                                ExecuteNonQuery(CenterLibrary.DBConnection.MyAdvantechGlobal, string.Format("insert into BB_ESTORE_ORDER values ('{0}','{1}','UnProcess', '', GETDATE(), GETDATE())", orderno, ERPID.ToString()));
                            else
                                ExecuteNonQuery(CenterLibrary.DBConnection.MyAdvantechGlobal, string.Format("insert into BB_ESTORE_ORDER values ('{0}','','NeedERPID', '', GETDATE(), GETDATE())", orderno));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string _smtp = ConfigurationManager.AppSettings.Get("MasterSMTP");
                string _from = ConfigurationManager.AppSettings.Get("MailFrom");
                string _to = ConfigurationManager.AppSettings.Get("MailTo");
                SmtpClient smtpClient1 = new SmtpClient(_smtp);
                MailMessage mail = new MailMessage(_from, _to);
                mail.Subject = "Check BB eStore order job failed";
                mail.IsBodyHtml = true;
                mail.Body = ex.ToString();
                try
                {
                    smtpClient1.Send(mail);
                }
                catch
                {

                }
            }
        }


        public static DataTable GetDataTable(string name, string sql)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = new SqlConnection(name);
            SqlDataAdapter da = new SqlDataAdapter(sql, conn);
            da.SelectCommand.CommandTimeout = 300;
            try
            {
                da.Fill(dt);
            }
            catch
            {
                if (conn.State != System.Data.ConnectionState.Closed)
                    conn.Close();
                if (conn != null)
                    conn.Dispose();
            }
            return dt;
        }

        public static Tuple<bool, string> ExecuteNonQuery(string name, string sql)
        {
            SqlConnection conn = new SqlConnection(name);
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 300;
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                return new Tuple<bool, string>(false, "Excute SQL error, SQL: " + sql + " Message + " + ex.ToString());
            }
            finally
            {
                if (conn.State != System.Data.ConnectionState.Closed)
                    conn.Close();
                conn.Dispose();
            }
            return new Tuple<bool, string>(true, string.Empty);
        }

        public static object ExecuteScalar(string ConnectionName, string strSqlCmd)
        {
            SqlConnection g_adoConn = new SqlConnection(ConnectionName);
            for (int i = 0; i <= 3; i++)
            {
                try
                {
                    g_adoConn.Open();
                    break; // TODO: might not be correct. Was : Exit For
                }
                catch (SqlException ex)
                {
                    if (i == 3)
                        throw ex;
                    System.Threading.Thread.Sleep(100);
                }
            }
            System.Data.SqlClient.SqlCommand dbCmd = g_adoConn.CreateCommand();
            dbCmd.CommandType = CommandType.Text;
            dbCmd.CommandText = strSqlCmd;
            dbCmd.CommandTimeout = 5 * 60;
            object retObj = null;
            try
            {
                retObj = dbCmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                g_adoConn.Close();
                throw ex;
            }
            g_adoConn.Close();
            return retObj;
        }
    }
}
