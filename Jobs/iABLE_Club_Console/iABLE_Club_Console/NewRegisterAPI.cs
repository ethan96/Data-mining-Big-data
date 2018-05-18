using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iABLE_Club_Console
{
    public class NewRegisterAPI
    {
        /// <summary>
        /// 每小時固定檢查 CurationActivity 註冊紀錄，目前只有需要幫 ATW 客戶添加點數
        /// </summary>
        public void RunRegister()
        {
            DataTable dt = SqlProvider.dbGetDataTable("MY", "select distinct ISNULL(a.EMAIL,''),ISNULL(a.COUNTRY,'') from CurationPool.dbo.CURATION_ACTIVITY_IMPORTED_LOG a where a.ACTIVITY_TYPE like '%Registration%' and a.SOURCE_TYPE in ('eStore','Corporate_Website') and COUNTRY in ('Taiwan','China') and a.TIMESTAMP between DATEADD(hour,-1,getdate()) and GETDATE() ");
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string userid = dr[0].ToString();
                    string country = dr[1].ToString();
                    if (!string.IsNullOrEmpty(userid) && !string.IsNullOrEmpty(country))
                    {
                        string storeID = string.Empty;
                        if (country.Equals("Taiwan", StringComparison.OrdinalIgnoreCase))
                            storeID = "ATW";
                        else if (country.Equals("China", StringComparison.OrdinalIgnoreCase))
                            storeID = "ACN";

                        if (!string.IsNullOrEmpty(storeID))
                        {
                            int result = 0;
                            object obj = SqlProvider.dbExecuteScalar("MY", string.Format(" select COUNT(*) from RewardRecord where UserID = '{0}'  and RewardID=0 and ActivityID=0 and TransactionType=0 and StoreID='{1}' ", userid.Trim(), storeID));
                            if (obj != null && int.TryParse(obj.ToString(), out result) == true && result == 0)
                            {
                                decimal point = 2;
                                if (storeID == "ACN")
                                {
                                    //2016/12/15-2017/3/15給4點，其餘都是2點
                                    if (DateTime.Compare(DateTime.Now, new DateTime(2016, 12, 15)) > 0 && DateTime.Compare(DateTime.Now, new DateTime(2017, 3, 15)) < 0)
                                        point = 4;
                                }
                                SqlProvider.dbExecuteNonQuery("MY", string.Format(" insert into RewardRecord (StoreID,UserID,RewardID,ActivityID,TransactionType,RecordType,OrderNo,Qty,Point,TotalPoint,CreatedBy,CreatedDate,SendMailStatus_Internal,SendMailStatus_Corporate,SendMailStatus_Sales) values ('{2}','{0}',0,0,0,1,'',1,{1},{1},'NewRegister',GETDATE(),1,1,1)", userid.Trim(), point, storeID));
                            }
                        }
                        
                    }
                }
            }
        }

    }
}
