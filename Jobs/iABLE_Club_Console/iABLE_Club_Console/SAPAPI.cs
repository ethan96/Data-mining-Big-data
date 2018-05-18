using iABLE_Club_Console.Model;
using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace iABLE_Club_Console
{
    public class SAPAPI
    {
        private StringBuilder _errMsg;

        public StringBuilder ErrMsg
        {
            get { return _errMsg; }
        }

        public SAPAPI()
        {
            this._errMsg = new StringBuilder();
        }

        public void CreateSAPOrder(RewardOrder rewardOrder)
        { 
            RewardService rewardService = new RewardService();

            SO_CREATE_COMMIT.SO_CREATE_COMMIT proxy1 = new SO_CREATE_COMMIT.SO_CREATE_COMMIT();
            SO_CREATE_COMMIT.BAPISDHD1 OrderHeader = new SO_CREATE_COMMIT.BAPISDHD1();
            SO_CREATE_COMMIT.BAPISDITMTable ItemIn = new SO_CREATE_COMMIT.BAPISDITMTable();
            SO_CREATE_COMMIT.BAPIPARNRTable PartNr = new SO_CREATE_COMMIT.BAPIPARNRTable();
            SO_CREATE_COMMIT.BAPISCHDLTable ScheLine = new SO_CREATE_COMMIT.BAPISCHDLTable();
            SO_CREATE_COMMIT.BAPICONDTable Conditions = new SO_CREATE_COMMIT.BAPICONDTable();
            SO_CREATE_COMMIT.BAPISDTEXTTable sdtextTable = new SO_CREATE_COMMIT.BAPISDTEXTTable();
            SO_CREATE_COMMIT.BAPIADDR1Table addr1Table = new SO_CREATE_COMMIT.BAPIADDR1Table();
            SO_CREATE_COMMIT.BAPIADDR1 PartnerAddressesRow = new SO_CREATE_COMMIT.BAPIADDR1();

            //OrderHeader Settings
            OrderHeader.Doc_Type = "ZOR";
            OrderHeader.Sales_Org = "TW01";
            OrderHeader.Distr_Chan = "10";
            OrderHeader.Division = "00";
            OrderHeader.Currency = "TWD";
            OrderHeader.Sales_Grp = "100";

            //Create address table
            PartnerAddressesRow.Name = rewardOrder.Receiver;
            PartnerAddressesRow.Country = "TW";
            PartnerAddressesRow.District = rewardOrder.Zip;
            PartnerAddressesRow.Distrct_No = rewardOrder.Zip;
            PartnerAddressesRow.E_Mail = rewardOrder.Email;
            PartnerAddressesRow.Street = rewardOrder.Zip + rewardOrder.address;
            PartnerAddressesRow.Tel1_Numbr = rewardOrder.Mobile;
            PartnerAddressesRow.Addr_No = "0000000001";
            addr1Table.Add(PartnerAddressesRow);

            //Create Ship to Partner
            SO_CREATE_COMMIT.BAPIPARNR PartNr_Ship_Record = new SO_CREATE_COMMIT.BAPIPARNR();
            PartNr_Ship_Record.Partn_Numb = "IOTMART"; //TODO ERP ID  要在SAP新增
            PartNr_Ship_Record.Partn_Role = "WE";
            PartNr_Ship_Record.Addr_Link = "0000000001";
            //PartNr_Ship_Record.Address = rewardOrder.address;
            //PartNr_Ship_Record.District = rewardOrder.Zip;
            //PartNr_Ship_Record.Telephone = rewardOrder.Mobile;
            //PartNr_Ship_Record.Telephone2 = rewardOrder.Tel;
            //PartNr_Ship_Record.Name = rewardOrder.Receiver;
            PartNr.Add(PartNr_Ship_Record);

            //Create Sold to Partner
            SO_CREATE_COMMIT.BAPIPARNR PartNr_Sold_Record = new SO_CREATE_COMMIT.BAPIPARNR();
            PartNr_Sold_Record.Partn_Role = "AG";
            PartNr_Sold_Record.Partn_Numb = "IOTMART"; //TODO ERP ID  要在SAP新增
            PartNr_Sold_Record.Addr_Link = "0000000001";
            //PartNr_Sold_Record.Address = rewardOrder.address;
            //PartNr_Sold_Record.District = rewardOrder.Zip;
            //PartNr_Sold_Record.Telephone = rewardOrder.Mobile;
            //PartNr_Sold_Record.Telephone2 = rewardOrder.Tel;
            //PartNr_Sold_Record.Name = rewardOrder.Receiver;
            PartNr.Add(PartNr_Sold_Record);

            //Create Sales Partner
            SO_CREATE_COMMIT.BAPIPARNR PartNr_Sales_Record = new SO_CREATE_COMMIT.BAPIPARNR();
            PartNr_Sales_Record.Partn_Role = "VE";
            PartNr_Sales_Record.Partn_Numb = "11000001"; //TC 提供的 Sales ID 
            PartNr.Add(PartNr_Sales_Record);

            //Item Record Settings
            //TODO 如果有多筆 Items, 請用 For each get Items
            int n = 1;
            foreach (var item in rewardOrder.items)
            {

                SO_CREATE_COMMIT.BAPISDITM Item_Record = new SO_CREATE_COMMIT.BAPISDITM();
                Item_Record.Material = rewardService.FindPartNoByRewardId((int)item.RewardID); //TODO PartNO
                Item_Record.Item_Categ = "ZTN3";
                Item_Record.Gross_Wght = (decimal)0.001;//要填多少??
                Item_Record.Untof_Wght = "KG";
                Item_Record.Itm_Number = n.ToString();//TODO 如果一張單有多個品項 Num + 1
                Item_Record.Ref_1 = "MyAdvantech";
                ItemIn.Add(Item_Record);

                //ScheLine Settings
                SO_CREATE_COMMIT.BAPISCHDL ScheLine_Record = new SO_CREATE_COMMIT.BAPISCHDL();
                ScheLine_Record.Itm_Number = Item_Record.Itm_Number;
                ScheLine_Record.Req_Qty = item.Qty.Value; //TODO 帶入 real qty             
                ScheLine_Record.Req_Date = DateTime.Now.AddDays(1).ToString("yyyyMMdd");//今天日期 +1
                ScheLine.Add(ScheLine_Record);

                //Condition Settings
                //SO_CREATE_COMMIT.BAPICOND Condition_Record = new SO_CREATE_COMMIT.BAPICOND();
                //Condition_Record.Itm_Number = Item_Record.Itm_Number;
                //Condition_Record.Tax_Code = "1";
                //Condition_Record.Access_Seq = "ZMWS";
                //Condition_Record.Cond_Type = "MWST";//TODO 禮品不帶價格
                //Condition_Record.Currency = "TWD";//TODO currency
                //Condition_Record.Cond_Value = 50000; //TODO total amount
                //Conditions.Add(Condition_Record);

                n++;
            }

            //Memo
            SO_CREATE_COMMIT.BAPISDTEXT S_HeaderTextsDt = new SO_CREATE_COMMIT.BAPISDTEXT();
            S_HeaderTextsDt.Doc_Number = rewardOrder.orderNo;//TODO 請填入 Order No
            S_HeaderTextsDt.Text_Id = "0002";//???
            S_HeaderTextsDt.Text_Line = "「iABLE 禮品兌換」, 加附感謝函";
            S_HeaderTextsDt.Langu = "TW";
            sdtextTable.Add(S_HeaderTextsDt);

            SO_CREATE_COMMIT.BAPISDTEXT S_HeaderTextsDt2 = new SO_CREATE_COMMIT.BAPISDTEXT();
            S_HeaderTextsDt2.Doc_Number = rewardOrder.orderNo;//TODO 請填入 Order No
            S_HeaderTextsDt2.Text_Id = "ZDH2";//???
            S_HeaderTextsDt2.Text_Line = "收件人: " + rewardOrder.Receiver + "\n" + "收件地址: " + rewardOrder.Zip + rewardOrder.address + "\n" +
                                         "電話: " + rewardOrder.Mobile + "\n" + "Email: " + rewardOrder.Email + "\n" + "快遞付費部門: " + rewardOrder.SalesDivCode + "\n" +
                                         "「iABLE 禮品兌換」, 加附感謝函";


            S_HeaderTextsDt2.Langu = "TW";
            sdtextTable.Add(S_HeaderTextsDt2);

            String SAPconnection = ConfigurationManager.AppSettings["isSAPConnTest"] == "true" ? "SAPConnTest" : "SAP_PRD"; //TODO true or false 可透過參數傳遞
            proxy1.Connection = new SAP.Connector.SAPConnection(ConfigurationManager.AppSettings[SAPconnection]);

            proxy1.Connection.Open();
            String strError = "", strRelationType = "", strPConvert = "", strpintnumassign = "";
            String strPTestRun = "";
            SO_CREATE_COMMIT.BAPIRET2Table retTable = new SO_CREATE_COMMIT.BAPIRET2Table();
            String refDoc_Number = rewardOrder.orderNo;//TODO 請填入 Order No

            SO_CREATE_COMMIT.BAPISDLS sdls = new SO_CREATE_COMMIT.BAPISDLS();
            SO_CREATE_COMMIT.BAPISDHD1X sdhd1x = new SO_CREATE_COMMIT.BAPISDHD1X();
            SO_CREATE_COMMIT.BAPI_SENDER sender = new SO_CREATE_COMMIT.BAPI_SENDER();
            SO_CREATE_COMMIT.BAPIPAREXTable parexTable = new SO_CREATE_COMMIT.BAPIPAREXTable();
            SO_CREATE_COMMIT.BAPICCARDTable ccardTable = new SO_CREATE_COMMIT.BAPICCARDTable();
            SO_CREATE_COMMIT.BAPICUBLBTable cublbTable = new SO_CREATE_COMMIT.BAPICUBLBTable();
            SO_CREATE_COMMIT.BAPICUINSTable cuinsTable = new SO_CREATE_COMMIT.BAPICUINSTable();
            SO_CREATE_COMMIT.BAPICUPRTTable cuprtTable = new SO_CREATE_COMMIT.BAPICUPRTTable();
            SO_CREATE_COMMIT.BAPICUCFGTable cucfgTable = new SO_CREATE_COMMIT.BAPICUCFGTable();
            SO_CREATE_COMMIT.BAPICUREFTable curefTable = new SO_CREATE_COMMIT.BAPICUREFTable();
            SO_CREATE_COMMIT.BAPICUVALTable cuvalTable = new SO_CREATE_COMMIT.BAPICUVALTable();
            SO_CREATE_COMMIT.BAPICUVKTable cuvktTable = new SO_CREATE_COMMIT.BAPICUVKTable();
            SO_CREATE_COMMIT.BAPICONDXTable condxTable = new SO_CREATE_COMMIT.BAPICONDXTable();
            SO_CREATE_COMMIT.BAPISDITMXTable sditmxTable = new SO_CREATE_COMMIT.BAPISDITMXTable();
            SO_CREATE_COMMIT.BAPISDKEYTable sdkeyTable = new SO_CREATE_COMMIT.BAPISDKEYTable();
            SO_CREATE_COMMIT.BAPISCHDLXTable sdhdlxTable = new SO_CREATE_COMMIT.BAPISCHDLXTable();

            proxy1.Bapi_Salesorder_Createfromdat2(strError, strRelationType, strPConvert, strpintnumassign, sdls,
                OrderHeader, sdhd1x, refDoc_Number, sender, strPTestRun, out refDoc_Number, ref parexTable, ref ccardTable,
                ref cublbTable, ref cuinsTable, ref cuprtTable, ref cucfgTable, ref curefTable, ref cuvalTable,
                ref cuvktTable, ref Conditions, ref condxTable, ref ItemIn, ref sditmxTable, ref sdkeyTable,
                ref PartNr, ref ScheLine, ref sdhdlxTable, ref sdtextTable, ref addr1Table, ref retTable);

            proxy1.CommitWork();
            proxy1.Connection.Close();

            // Wait a while.
            Thread.Sleep(5000);

            ////檢查Create SO 成功 --> unblock GP(because of ZTN3) --> 建DN
            int count = 1;
            foreach (SO_CREATE_COMMIT.BAPIRET2 ret in retTable)
                if (!ret.Type.Equals("S"))
                {
                    rewardOrder.SoMsg += "(" + count.ToString() + ")" + ret.Message;
                    count += 1;
                }
            if (rewardOrder.SoMsg.Length == 0)
                rewardOrder.SoMsg = string.Format("Success");

            if (rewardOrder.SoMsg == "Success")
            {


                // unblock GP
                String SAPOracleConnection = ConfigurationManager.AppSettings["isSAPConnTest"] == "true" ? "SAP_Test" : "SAP_PRD"; //TODO true or false 可透過參數傳遞
                int output = 0;
                string sqlSOGPBlockLines =
                    " select POSNR, LSSTA from saprdp.vbup where LSSTA='C' and vbeln='" + rewardOrder.orderNo + "' ";

                DataTable dtSOGPLines = OraDBUtil.dbGetDataTable(SAPOracleConnection, sqlSOGPBlockLines);


                if (dtSOGPLines.Rows.Count > 0)
                {
                    Z_RELEASE_GP_ITEM.Z_RELEASE_GP_ITEM pro1 = new Z_RELEASE_GP_ITEM.Z_RELEASE_GP_ITEM(ConfigurationManager.AppSettings[SAPconnection]);
                    pro1.Connection.Open();
                    foreach (DataRow GPLineRow in dtSOGPLines.Rows)
                    {
                        pro1.Z_Release_Gp_Item(GPLineRow["POSNR"].ToString(), rewardOrder.orderNo, "", out output);
                    }
                    pro1.Connection.Close();
                }



                Create_DN.BAPIDLVREFTOSALESORDERTable salesorderTable = new Create_DN.BAPIDLVREFTOSALESORDERTable();
                int line_no = 1;
                foreach (var item in rewardOrder.items)
                {
                    Create_DN.BAPIDLVREFTOSALESORDER salesOrder = new Create_DN.BAPIDLVREFTOSALESORDER();
                    salesOrder.Ref_Doc = rewardOrder.orderNo;
                    salesOrder.Ref_Item = line_no.ToString();
                    salesOrder.Dlv_Qty = (decimal)item.Qty;
                    salesorderTable.Add(salesOrder);
                    line_no++;
                }


                Create_DN.Create_DN dn = new Create_DN.Create_DN();
                dn.Connection = new SAP.Connector.SAPConnection(ConfigurationManager.AppSettings[SAPconnection]);
                dn.Connection.Open();

                Create_DN.BAPIDLVITEMCREATEDTable dlvItemCreatedTable = new Create_DN.BAPIDLVITEMCREATEDTable();
                Create_DN.BAPISHPDELIVNUMBTable shpDelivNumbTable = new Create_DN.BAPISHPDELIVNUMBTable();
                Create_DN.BAPISHPDELIVNUMB shpDelivNumb = new Create_DN.BAPISHPDELIVNUMB();


                Create_DN.BAPIPAREXTable parexTableIn = new Create_DN.BAPIPAREXTable();
                Create_DN.BAPIPAREX parex = new Create_DN.BAPIPAREX();



                Create_DN.BAPIPAREXTable parexTableOut = new Create_DN.BAPIPAREXTable();
                Create_DN.BAPIRET2Table ret2Table = new Create_DN.BAPIRET2Table();
                Create_DN.BAPIDLVSERIALNUMBERTable dlvSerialNumberTable = new Create_DN.BAPIDLVSERIALNUMBERTable();
                Create_DN.BAPIDLVREFTOSTOTable dlvRefToStoTable = new Create_DN.BAPIDLVREFTOSTOTable();
                Create_DN.BAPIDLVREFTOSTO dlvrefto = new Create_DN.BAPIDLVREFTOSTO();


                dn.Z_Create_Delivery_Note("1", ref dlvItemCreatedTable, ref shpDelivNumbTable, ref parexTableIn,
                    ref parexTableOut, ref ret2Table, ref salesorderTable, ref dlvSerialNumberTable, ref dlvRefToStoTable);
                dn.CommitWork();
                dn.Connection.Close();
                Thread.Sleep(5000);
                //在 SALES_ORDER_ITEMS 傳入REF_DOC(Sales Order), REF_ITEM, DLV_QTY 即可

                //檢查Create DN 有無成功
                count = 1;
                string dnNo = "";
                foreach (Create_DN.BAPIRET2 ret in ret2Table)
                {
                    if (!ret.Type.Equals("S"))
                    {
                        rewardOrder.DnMsg += "(" + count.ToString() + ")" + ret.Message;
                        count += 1;
                    }
                    if (ret.Id == "BAPI")
                        dnNo = ret.Message_V2;
                }
                if (rewardOrder.DnMsg.Length == 0)
                {
                    rewardOrder.DN = dnNo;
                    rewardOrder.DnMsg = string.Format("Success");
                }  
            }




        }
    }
}
