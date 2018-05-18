using iABLE_Club_Console.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Configuration;
using System.Data;

namespace iABLE_Club_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            /*---------------------------以下 job 會每小時進行------------------------------------------------------*/

            //Step 1. 每小時檢查註冊紀錄
            ProcessNewRegister();

            //Step 2. 每小時檢查有無客戶兌換禮品，若有則發通知信給內部 & Iris
            ProcessSendInternalEmail();

            //Step 3. 每小時檢查有無corp. 手動出貨紀錄，有手動出貨則要自動建立 Siebel Activity
            CreateSiebelActivityFromManuallyRecords();

            /*---------------------------以下 job 不會每小時進行------------------------------------------------------*/

            //Step 1. 每天下午 5 點，a.會自動找已經轉DN的單，若SAP狀態為出貨，則UPDATE回IABLE出貨狀態為已出貨
            //                      b.item皆已出貨的訂單，發信給Sales
            //                      c.轉成倉 DN，凡是有料號的兌換紀錄，皆會自動轉成倉開 DN 出貨
            int hr = DateTime.Now.Hour;
            string saporder = ConfigurationManager.AppSettings["SAPorder"] ?? "N";
            RewardService rewardService = new RewardService();

            if (string.Equals(saporder.ToUpper(), "Y") && hr == 17)
            {
                //確認已轉DN的單狀態是否為已出貨，是的話就Update回IABLE出貨狀態
                List<RewardRecord> readyToShipRecords = rewardService.GetReadyToShipRecords();                
                if (readyToShipRecords != null)
                     rewardService.UpdateTranTypeForReadyToShipRecords(readyToShipRecords);

                //找出未發尚未發信給Sales的訂單，若訂單中的items DN狀態都為已出貨，則寄信通知sales並Create Siebel Activity
                List<RewardOrder> readyToSendSalesMailOrders = rewardService.GetUnSendSalesMailOrders().Where(o => o.DnMsg != "Contains Unshipped DN").ToList();

                if (readyToSendSalesMailOrders.Count > 0 && ConfigurationManager.AppSettings["isSAPConnTest"] != "true")
                {
                    ProcessSendSalesAndCorpEmail(readyToSendSalesMailOrders);
                    CreateSiebelActivity(readyToSendSalesMailOrders);
                }

                //成倉轉DN
                List<RewardOrder> orders = ProcessSAPOrder();           

            }

            //Step 2. 每月 15 日早上 8 點，會檢查兌換紀錄，凡是好書好禮類，並且沒有料號的紀錄，就會發送 email 給 corp. 請他們手動寄出禮品
            int day = DateTime.Now.Day;
            if (day == 15 && hr == 8)
                SendMailForCorpEvery15thEachMonth();

        }

        public static List<RewardOrder> ProcessSAPOrder()
        {
            List<string> errMsgBoxs = new List<string>();
            //List<string> successOrderMsgBoxs = new List<string>();
            Email mail = new Email();
            RewardService rewardService = new RewardService();
            IEnumerable<IGrouping<string, RewardRecord>> recordGroup = null;
            List<RewardOrder> rewardOrders = new List<RewardOrder>();
            try
            {
                List<RewardRecord> rewardRecords = rewardService.GetUnOrderedRecords();

                if (rewardRecords != null)
                {
                    recordGroup = rewardRecords.GroupBy(r => r.OrderNo);
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
                        order.DnMsg = "";
                        order.DN = "";
                        order.StoreID = firstItem.StoreID;
                        order.RecordID = firstItem.RecordID;
                        order.SalesEmail = firstItem.SalesEmail;
                        order.SalesDivCode = rewardService.FindSalesDivCodeBySalesEmail(firstItem); //找尋對應sales的 Cost center
                        if (!String.IsNullOrEmpty(order.SalesDivCode))
                        {
                            foreach (var item in g)
                                order.items.Add(item);
                            rewardOrders.Add(order);
                        }
                        else
                            order.SoMsg = string.Format("Can not find Sales Div Code(Cost center) in {0} is not correct.", order.orderNo);
                    }

                    if (rewardOrders != null)
                    {
                        foreach (var o in rewardOrders)
                        {
                            if (o.orderNo.StartsWith("IABLE"))
                            {

                                if (o.Receiver == null || o.Email == null || o.Zip == null || o.address == null || o.Mobile == null)
                                    o.SoMsg = string.Format("The order: {0} has imcomplete contact information.", o.orderNo);
                                else
                                {
                                    SAPAPI sappapi = new SAPAPI();
                                    sappapi.CreateSAPOrder(o);
                                    if (o.SoMsg == "Success" && o.DnMsg == "Success")
                                    {
                                        foreach (var item in o.items)
                                            rewardService.UpdateRewardRecordTranscationType(item, 5);

                                    }
                                }
                            }
                            else
                            {
                                o.SoMsg = string.Format("The format of Order NO: {0} is not correct.", o.orderNo);
                            }
                       }
                    }
                }
            }
            catch (Exception ex)
            {
                errMsgBoxs.Add(ex.Message);
            }


            #region Mail message to IT and Create Log
            SendMailForSOandDNStatus(rewardOrders, errMsgBoxs);

            // 暫時寄信給OP拋單到倉庫列印已開出的DN
            SendMailToOPForPrintDN(rewardOrders.Where(o=>string.IsNullOrEmpty(o.DN) == false).ToList());
            CreateLog(errMsgBoxs);

            #endregion

            return rewardOrders;
        }


        public static void SendMailForSOandDNStatus(List<RewardOrder> orders, List<string> errMsgBoxs)
        {
            System.Web.UI.WebControls.GridView gv = new System.Web.UI.WebControls.GridView();
            gv.DataSource = orders.Select(p => new { OrderNo = p.orderNo, SOMessage = p.SoMsg, DNMessage = p.DnMsg, DN = p.DN }).ToList();
            gv.DataBind();
            StringBuilder sb = new StringBuilder();
            System.IO.StringWriter sw = new System.IO.StringWriter(sb);
            System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(sw);
            gv.RenderControl(hw);

            Email email = new Email();
            email.Subject =  string.Format("Create IABLE SO and DN in {0}!",  DateTime.Now.ToString("yyyy-MM-dd")); ;
            email.MailToAddress = email.MailToAddress;
            email.MailBody = string.Format("<h2><font color =\"blue\">{0}</font> Reward Orders need to be transfered to SAP and create delivery note today:</h2>", orders.Count.ToString());
            email.MailBody += sb.ToString();


            if (errMsgBoxs!=null && errMsgBoxs.Count > 0)
            {
                email.MailBody += "<br />" + "<strong>Some technical Errors happen today:</strong>" + "<br />";
                foreach (var m in errMsgBoxs)
                {
                    email.MailBody += m;
                    email.MailBody += "<br />" + "--------------------------------------------" + "<br />";
                }
            }

            email.SendEmail();
        }

        public static void SendMailToOPForPrintDN(List<RewardOrder> orders)
        {
            if (orders.Count > 0)
            {
                System.Web.UI.WebControls.GridView gv = new System.Web.UI.WebControls.GridView();
                gv.DataSource = orders.Select(p => new { OrderNo = p.orderNo, DN = p.DN }).ToList();
                gv.DataBind();
                StringBuilder sb = new StringBuilder();
                System.IO.StringWriter sw = new System.IO.StringWriter(sb);
                System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(sw);
                gv.RenderControl(hw);

                Email email = new Email();
                email.Subject = string.Format("(iABLE)有新iABLE禮品兌換DN單成立，請幫忙進行拋單!"); ;
                email.MailToAddress = ConfigurationManager.AppSettings.Get("MailToOP");
                email.CC = ConfigurationManager.AppSettings.Get("MailTo") + ",Wen.Chiang@advantech.com.tw";
                email.MailBody = string.Format("Dear Tiffany,Joan 您好，<br /><br />以下iABLE禮品兌換訂單已開立DN，請幫忙將DN拋單至倉庫印出，以利後續出貨事宜，謝謝!!<br /><br /> ");
                email.MailBody += sb.ToString();
                email.SendEmail();
            }
            
        }

        public static void CreateLog(List<string> errMsgBoxs)
        {
            RewardService rewardService = new RewardService();

            if (errMsgBoxs != null && errMsgBoxs.Count > 0)
            {
                RewardUserLog log = new RewardUserLog();
                log.StoreID = "ATW";
                log.UserID = "iAbleClub_Console";
                log.LogType = "Transfer to SAP and Create DN Error";
                foreach (var m in errMsgBoxs)
                    log.Message += m;
                log.CreatedBy = "iAbleClub_Console";
                log.CreatedDate = DateTime.Now;
                try
                {
                    rewardService.CreateUserLog(log);
                }
                catch
                {

                }              
            }

        }

        public static void CreateSiebelActivity(List<RewardOrder> orders)
        {
            if (orders != null && orders.Count > 0)
            {
                SiebelAPI siebel = new SiebelAPI();
                
                foreach (RewardOrder order in orders)
                    siebel.CreateSiebleActivity(order);

                SendMail(orders, string.Format("Create IABLE Rewards to Siebel Activity in {0}", DateTime.Now.ToString("yyyy-MM-dd")));
            }
        }

        public static void ProcessNewRegister()
        {
            NewRegisterAPI regi = new NewRegisterAPI();
            try
            {
                regi.RunRegister();
            }
            catch (Exception ex)
            {
                SmtpClient smtpClient1 = new SmtpClient(ConfigurationManager.AppSettings["MasterSMTP"]);
                smtpClient1.Send(ConfigurationManager.AppSettings.Get("MailFrom"), ConfigurationManager.AppSettings.Get("MailTo"), "iABLE create new register failed", ex.ToString());
            }
            finally
            {
                regi = null;
            }
            
        }

        public static void ProcessSendInternalEmail()
        {
            //Check data first
            object count = SqlProvider.dbExecuteScalar("MY", " SELECT COUNT(*) FROM RewardRecord where RecordType = 2 AND SendMailStatus_Internal = 0 ");
            if (count != null && Convert.ToInt32(count) > 0)
            {
                iAbleClubService.IMartPointService ws = new iAbleClubService.IMartPointService();
                try
                {
                    ws.Timeout = 60000;
                    ws.SendMailForInternal(ConfigurationManager.AppSettings.Get("WebServicePassword"));
                }
                catch (Exception ex)
                {
                    SmtpClient smtpClient1 = new SmtpClient(ConfigurationManager.AppSettings["MasterSMTP"]);
                    smtpClient1.Send(ConfigurationManager.AppSettings.Get("MailFrom"), ConfigurationManager.AppSettings.Get("MailTo"), "iABLE send internal mail failed", ex.ToString());
                }
                finally
                {
                    ws.Dispose();
                    ws = null;
                }
            }
        }

        public static void ProcessSendSalesAndCorpEmail(List<RewardOrder> orders)
        {
            if (orders != null && orders.Count > 0)
            {
                //List<string> success = orders.Where(o => o.SoMsg == "Success" && o.DnMsg == "Sucess").Select(p => p.orderNo).ToList();
                //if (success.Count > 0)
                //{
                    iAbleClubService.IMartPointService ws = new iAbleClubService.IMartPointService();
                    try
                    {
                        ws.Timeout = 60000;
                        ws.SendMailForSalesAndCorpAfterSAPorderSuccess(ConfigurationManager.AppSettings.Get("WebServicePassword"), orders.Select(p => p.orderNo).ToArray());
                    }
                    catch (Exception ex)
                    {
                        SmtpClient smtpClient1 = new SmtpClient(ConfigurationManager.AppSettings["MasterSMTP"]);
                        smtpClient1.Send(ConfigurationManager.AppSettings.Get("MailFrom"), ConfigurationManager.AppSettings.Get("MailTo"), "iABLE send sales and corp mail failed", ex.ToString());
                    }
                    finally
                    {
                        ws.Dispose();
                        ws = null;
                    }

                //}
            }

        }

        public static void CreateSiebelActivityFromManuallyRecords()
        {
            try
            {
                System.Data.DataTable dt = SqlProvider.dbGetDataTable("MY", " select * from RewardRecord where RewardID > 0 and TransactionType = 2 and (Tel2 is null or Tel2='ERROR') ");
                if (dt != null && dt.Rows.Count > 0)
                {
                    SiebelAPI siebel = new SiebelAPI();
                    List<RewardOrder> orders = new List<RewardOrder>();
                    foreach (System.Data.DataRow dr in dt.Rows)
                    {
                        RewardOrder order = new RewardOrder();
                        order.orderNo = dr["OrderNo"].ToString();
                        order.Email = dr["EmailAddress1"].ToString();
                        order.SalesEmail = dr["SalesEmail"].ToString();
                        order.StoreID = dr["StoreID"].ToString();
                        order.RecordID = int.Parse(dr["RecordID"].ToString());
                        siebel.CreateSiebleActivity(order);
                        orders.Add(order);
                    }

                    SendMail(orders, string.Format("Create IABLE Rewards to Siebel Activity from manually records in {0}", DateTime.Now.ToString("yyyy-MM-dd")));
                }
            }
            catch (Exception ex)
            {
                SmtpClient smtpClient1 = new SmtpClient(ConfigurationManager.AppSettings["MasterSMTP"]);
                smtpClient1.Send(ConfigurationManager.AppSettings.Get("MailFrom"), ConfigurationManager.AppSettings.Get("MailTo"), "iABLE create Siebel Activity from manually records failed", ex.ToString());
            }

        }

        public static void SendMailForCorpEvery15thEachMonth()
        {
            iAbleClubService.IMartPointService ws = new iAbleClubService.IMartPointService();
            try
            {
                ws.Timeout = 60000;
                ws.SendMailForCorpEvery15thEachMonth();
            }
            catch (Exception ex)
            {
                SmtpClient smtpClient1 = new SmtpClient(ConfigurationManager.AppSettings["MasterSMTP"]);
                smtpClient1.Send(ConfigurationManager.AppSettings.Get("MailFrom"), ConfigurationManager.AppSettings.Get("MailTo"), "iABLE send corp mail failed", ex.ToString());
            }
            finally
            {
                ws.Dispose();
                ws = null;
            }
        }

        public static void SendMail(List<RewardOrder> orders, string subject)
        {
            System.Web.UI.WebControls.GridView gv = new System.Web.UI.WebControls.GridView();
            gv.DataSource = orders.Select(p => new { OrderNo = p.orderNo, UserID = p.Email, Status = p.SiebelActivityStatus.ToString(), RowIDOrMessage = p.SiebelActivityMessage }).ToList();
            gv.DataBind();
            StringBuilder sb = new StringBuilder();
            System.IO.StringWriter sw = new System.IO.StringWriter(sb);
            System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(sw);
            gv.RenderControl(hw);

            Email email = new Email();
            email.Subject = subject;
            email.MailToAddress = email.MailToAddress + ",Iris.Tsai@advantech.com.tw";
            email.MailBody = sb.ToString();
            email.SendEmail();
        }

    }
}
