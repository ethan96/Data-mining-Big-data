using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iABLE_Club_Console
{
    public class SiebelAPI
    {
        public void CreateSiebleActivity(RewardOrder order)
        {
            if (order.RecordID == null)
            {
                order.SiebelActivityStatus = false;
                order.SiebelActivityMessage = "No RecordID";
                return;
            }

            if (string.IsNullOrEmpty(order.Email))
            {
                SqlProvider.dbExecuteNonQuery("MY", string.Format("update RewardRecord set Tel2='ERROR' where RecordID= {0} ", order.RecordID));
                order.SiebelActivityStatus = false;
                order.SiebelActivityMessage = "No customer email";
                return;
            }

            if (string.IsNullOrEmpty(order.SalesEmail))
            {
                SqlProvider.dbExecuteNonQuery("MY", string.Format("update RewardRecord set Tel2='ERROR' where RecordID= {0} ", order.RecordID));
                order.SiebelActivityStatus = false;
                order.SiebelActivityMessage = "No sales email";
                return;
            }

            if (string.IsNullOrEmpty(order.StoreID))
            {
                SqlProvider.dbExecuteNonQuery("MY", string.Format("update RewardRecord set Tel2='ERROR' where RecordID= {0} ", order.RecordID));
                order.SiebelActivityStatus = false;
                order.SiebelActivityMessage = "No StoreID";
                return;
            }

            var addAction = new Siebel_AddAction.ADVWebServoceAddAction();
            var action = new Siebel_AddAction.ACT();
            var actionInput = new Siebel_AddAction.AddAction_Input();
            var actionOutput = new Siebel_AddAction.AddAction_Output();

            action.ACT_TYPE = "Email - Outbound";
            action.COMMENT = "";
            action.DESP = "客戶於IABLE兌換禮品";
            action.STATUS = "Not Started";
            //action.CON_ROW_ID = "";
            action.CONTACT_EMAIL = order.Email.Trim();
            action.ORG = order.StoreID;
            action.SALES_LEADS_FLAG = "Y";
            action.OWNER_EMAIL = order.SalesEmail.Split(',')[0]; //只取第一個 sales 當作 primary sales
            action.PLANNED_START = DateTime.Now.ToString("MM/dd/yyyy");
            action.SRC_ROW_ID = "1-19E0SVD";//Add activity source

            actionInput.ACT = action;
            actionInput.SOURCE = "MTL";//Change to MTL

            try
            {
                actionOutput = addAction.AddAction(actionInput);
            }
            catch (Exception ex)
            {
                SqlProvider.dbExecuteNonQuery("MY", string.Format("update RewardRecord set Tel2='ERROR' where RecordID= {0} ", order.RecordID));
                order.SiebelActivityStatus = false;
                order.SiebelActivityMessage = ex.ToString();
            }

            if (actionOutput.STATUS.ToUpper() == "SUCCESS")
            {
                SqlProvider.dbExecuteNonQuery("MY", string.Format("update RewardRecord set Tel2='{0}' where RecordID= {1} ", actionOutput.ROW_ID, order.RecordID));
                order.SiebelActivityStatus = true;
                order.SiebelActivityMessage = actionOutput.ROW_ID;
            }
            else
            {
                SqlProvider.dbExecuteNonQuery("MY", string.Format("update RewardRecord set Tel2='ERROR' where RecordID= {0} ", order.RecordID));
                order.SiebelActivityStatus = false;
                order.SiebelActivityMessage = string.Format("Error code: {0}, Error message: {1}", actionOutput.Error_spcCode, actionOutput.Error_spcMessage);
            }
        }

    }
}
