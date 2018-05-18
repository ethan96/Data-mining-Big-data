using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace iABLE_Club_Console.Model
{
    public class RewardService
    {
        private  MyAdvantechGlobalEntities _entity { get; set; }

        public RewardService()
        {
            _entity = new MyAdvantechGlobalEntities();
        }

        public List<RewardRecord> GetUnOrderedRecords()
        {
            int transactionType = ConfigurationManager.AppSettings["isSAPConnTest"] == "true" ? 0 : 1;
            List<RewardRecord> query    //如果是要轉SAP測試站，則只取測試訂單來轉(TransactionType = 0 && RecordType = 2) , 目前只取好禮來轉SAP
                = _entity.RewardRecord.Where(r => r.TransactionType == transactionType && r.RecordType == 2 && r.ActivityID == 3).ToList();
            return query;
        }

        public List<RewardRecord> GetReadyToShipRecords()
        {
            List<RewardRecord> query    
                = _entity.RewardRecord.Where(r => r.TransactionType == 5 && r.RecordType == 2 && r.ActivityID == 3).ToList();
            return query;
        }

        public List<RewardOrder> GetUnSendSalesMailOrders()
        {
            IEnumerable<IGrouping<string, RewardRecord>> recordGroup = null;

            List<RewardRecord> UnSendSalesMailRecords
                = _entity.RewardRecord.Where(r => r.SendMailStatus_Sales == 0).ToList();

            List<RewardOrder> rewardOrders = new List<RewardOrder>();
            if (UnSendSalesMailRecords != null)
            {
                recordGroup = UnSendSalesMailRecords.GroupBy(r => r.OrderNo);
                foreach (var g in recordGroup)
                {

                    RewardOrder order = new RewardOrder();
                    order.orderNo = g.Key;
                    var firstItem = g.FirstOrDefault();
                    order.Receiver = firstItem.Receiver1;
                    order.Email = firstItem.EmailAddress1;
                    order.Zip = firstItem.Zip1;
                    order.address = firstItem.Address1;
                    order.Mobile = firstItem.Mobile1;
                    order.Tel = firstItem.Tel1;
                    order.CompanyName = firstItem.CompanyName1;
                    order.SalesEmail = firstItem.SalesEmail;
                    order.SoMsg = "";                    
                    order.StoreID = firstItem.StoreID;
                    order.RecordID = firstItem.RecordID;
                    order.SalesEmail = firstItem.SalesEmail;

                    foreach (var item in g)
                    {
                        order.items.Add(item);
                        if(item.TransactionType !=2)
                            order.DnMsg = "Contains Unshipped DN";
                    }
                    rewardOrders.Add(order);
                }
            }
            return rewardOrders;
        }

        public void CreateUserLog(RewardUserLog log)
        {
            try
            {
                _entity.RewardUserLog.Add(log);
                _entity.SaveChanges();
            }
            catch (DbEntityValidationException e)
            {
                foreach (var eve in e.EntityValidationErrors)
                {
                    Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                        eve.Entry.Entity.GetType().Name, eve.Entry.State);
                    foreach (var ve in eve.ValidationErrors)
                    {
                        Console.WriteLine("- Property: \"{0}\", Error: \"{1}\"",
                            ve.PropertyName, ve.ErrorMessage);
                    }
                }
                throw;
            }

        }

        public void UpdateRewardRecordTranscationType(RewardRecord record,int type)
        {
            var result = _entity.RewardRecord.SingleOrDefault(r => r.RecordID == record.RecordID);
            if (result != null)
            {
                result.TransactionType = type;
                _entity.SaveChanges();
            }
        }

        public string FindPartNoByRewardId(int id)
        {
            return _entity.RewardItem.Where(r => r.RewardID == id).FirstOrDefault().PartNo;
        }

        public string FindSalesDivCodeBySalesEmail(RewardRecord record)
        {
            string userid = "";
            try
            {

                DataTable dt = SqlProvider.dbGetDataTable("EZ", string.Format(" SELECT DISTINCT a.EMAIL_ADDR, prof.DIV_NAME, prof.SHIFT_NAME, prof.ORG_DESC, prof.REGION,prof.JobFamilyNameEnglish, prof.DIV_CODE FROM  EZ_PROFILE a  LEFT OUTER JOIN VW_PROFILE prof ON a.EZROWID = prof.EZROWID WHERE (a.EMAIL_ADDR = '{0}') ORDER BY a.EMAIL_ADDR", record.SalesEmail));
                if (dt != null && dt.Rows.Count > 0)
                {
                    userid = dt.Rows[0]["DIV_CODE"].ToString();
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
            return userid;
        }


        public void UpdateTranTypeForReadyToShipRecords(List<RewardRecord> readyToShipRecords)
        {
            DataTable dt = new DataTable();
            try
            {
                foreach (var record in readyToShipRecords)
                {
                    dt = OraDBUtil.dbGetDataTable("SAP_PRD", string.Format("SELECT c.VBELN,a.VBELN,a.POSNR,a.MATNR," +
                                                                            "(a.LFIMG - NVL((SELECT SUM(CASE WHEN b.VBTYP_N = 'R' THEN b.RFMNG * 1 ELSE b.RFMNG * -1 END) FROM SAPRDP.VBFA b WHERE b.VBELV = a.VBELN AND b.POSNV = a.POSNR AND b.VBTYP_N  in ('R', 'H', 'h') group by b.VBELV, b.POSNV), 0)) as Openqty " +
                                                                            " FROM SAPRDP.LIPS a INNER JOIN SAPRDP.VBAP b ON b.VBELN = a.VGBEL AND b.POSNR = a.VGPOS  INNER JOIN SAPRDP.VBAK c ON c.VBELN = b.VBELN " +
                                                                            "WHERE a.MANDT = '168' " +
                                                                            "AND c.VBELN = '{0}' " +
                                                                            "AND a.MATNR = '{1}'",record.OrderNo, FindPartNoByRewardId((int)record.RewardID)));
                    if (dt != null && dt.Rows.Count > 0 && dt.Rows[0]["OPENQTY"].ToString() == "0")
                    {
                        UpdateRewardRecordTranscationType(record, 2);
                    }
                }
                
            }

            catch (Exception ex)
            {
                SmtpClient smtpClient1 = new SmtpClient(ConfigurationManager.AppSettings["MasterSMTP"]);
                smtpClient1.Send(ConfigurationManager.AppSettings.Get("MailFrom"), ConfigurationManager.AppSettings.Get("MailTo"), "iABLE update Shipment status fail!", ex.ToString());
            }
        }
    }
}
