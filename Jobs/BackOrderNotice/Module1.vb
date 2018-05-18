Imports System.Text
Imports System.Reflection
Imports System.IO

Module Module1

    Sub Main()
        GetBackOrder()
    End Sub

    Private Function GetBackOrder() As DataTable
        Dim Weeksdays() As String = {"MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"}
        Dim weekint As Integer = Now.DayOfWeek() - 1
        Dim ERPSQL As New StringBuilder()
        ERPSQL.AppendFormat("   Select  STUFF ( (  ")
        ERPSQL.AppendFormat(" Select distinct ','''+ COMPANYID  +'''' FROM   ScheduleHead   ")
        ERPSQL.AppendFormat("WHERE ")
        ERPSQL.AppendFormat(" {0} = 1 ", Weeksdays(weekint))
        ERPSQL.AppendFormat(" for xml path('')   ")
        ERPSQL.AppendFormat("  ),1,1,'') AS  ERPLIST  ")
        Dim Erpids As Object = dbUtil.dbExecuteScalar(CenterLibrary.DBConnection.MyLocal, ERPSQL.ToString())
        If Erpids Is Nothing OrElse String.IsNullOrEmpty(Erpids) Then
            dbUtil.SendEmailWithAttachment("frank.chung@advantech.com.tw", "myadvantech@advantech.com", "BackOrder(Mails) send :  has no corresponding CompanyIDs ", "", False, "", "", Nothing, "")
            Exit Function
        End If
        Dim sb As New StringBuilder()
        sb.AppendFormat("   select distinct  T.SoldTo ,T.BILLTOID,T.COMPANYNAME FROM  ")
        sb.AppendFormat("   ( select  ISNULL( C.PARENTCOMPANYID,AB.BILLTOID) AS SoldTo ,AB.BILLTOID,C.COMPANY_NAME as COMPANYNAME ")
        sb.AppendFormat("    from  SAP_BACKORDER_AB AB LEFT JOIN  SAP_DIMCOMPANY  C  ")
        sb.AppendFormat("  ON AB.BILLTOID = C.COMPANY_ID  where C.ORG_ID ='EU10' and AB.SHIPTOID <> '' and AB.SHIPTOID is not null  ")
        sb.AppendFormat("   ) AS T ")
        sb.AppendFormat("   WHERE  T.SoldTo IN ({0})  ", Erpids)
        Dim dt As DataTable = dbUtil.dbGetDataTable(CenterLibrary.DBConnection.MyAdvantechGlobal, sb.ToString())
        Dim dtDetai As New DataTable
        For i As Integer = 0 To dt.Rows.Count - 1
            Try
                dtDetai.Clear()
                dtDetai = GetBackOrderDetail(dt.Rows(i)("SoldTo"))
                Console.Write(dt.Rows(i)("SoldTo"))
                Console.WriteLine()
                Dim exePath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim strFPath As String = exePath + String.Format("\{0}.xls", dt.Rows(i)("SoldTo"))
                If IO.File.Exists(strFPath) = True Then
                    IO.File.Delete(strFPath)
                End If
                dbUtil.DataTable2ExcelFile(dtDetai, strFPath)
                Dim fileStream As New FileStream(strFPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim bytes As Byte() = New Byte(fileStream.Length - 1) {}
                fileStream.Read(bytes, 0, bytes.Length)
                fileStream.Close()
                Dim stream As Stream = New MemoryStream(bytes)
                Dim Subject As String = String.Format("BackOrder for {0}({1})", dt.Rows(i)("COMPANYNAME"), dt.Rows(i)("SoldTo"))
                'Dim dr() As DataRow = CompanyList.Select(String.Format("CompanyID='{0}'", dt.Rows(i)("SoldTo")))
                Dim salesmails As String = ""


                Dim sql As New StringBuilder
                With sql
                    .AppendFormat("  Select  STUFF ( ( ")
                    .AppendFormat("Select distinct ';'+ MAIL  +''  from [dbo].[ScheduleMails] where HeadID =(  select  top 1 ID   from  ScheduleHead  where COMPANYID = '{0}') ", dt.Rows(i)("SoldTo"))
                    .AppendFormat(" for xml path('')  ),1,1,'') AS  ERPLIST  ")
                End With

                Dim mails As Object = dbUtil.dbExecuteScalar(CenterLibrary.DBConnection.MyLocal, sql.ToString())


                'If dr.Length > 0 Then
                '    If Not IsDBNull(dr(0).Item("Tomails")) Then
                '        salesmails = dr(0).Item("Tomails").ToString().Trim().Trim(New Char() {";"})
                '    End If
                '    'If Not IsDBNull(dr(0).Item("IsCallSAP")) AndAlso dr(0).Item("IsCallSAP") = "1" Then
                '    '    salesmails = salesmails + ";" + GetSalesMails(dt.Rows(i)("SoldTo"))
                '    'End If
                'End If
                If mails IsNot Nothing AndAlso Not String.IsNullOrEmpty(mails) Then
                    salesmails = mails.ToString().Trim().Trim(New Char() {";"})
                End If
                ' salesmails = salesmails.ToString().Trim().Trim(New Char() {";"})
                'Dim body As String = String.Format("This is a test for IT, it will be sent to "" {0} """, salesmails)

                Dim body As String = String.Format("")

                'Frank 20160426 send test email to Frank and Ryan
                'dbUtil.SendEmailWithAttachment("frank.chung@advantech.com.tw;fan245@gmail.com", "myadvantech@advantech.com", Subject, body, False, "", "", stream, String.Format("{0}.xls", dt.Rows(i)("SoldTo")))


                'Frank 20160426 below line will send the mail to real user and CP customers
                dbUtil.SendEmailWithAttachment(salesmails, "myadvantech@advantech.com", Subject, body, False, "", "", stream, String.Format("{0}.xls", dt.Rows(i)("SoldTo")))

            Catch ex As Exception
                Console.Write(ex.Message.ToString)
                Console.WriteLine()
                dbUtil.SendEmailWithAttachment("frank.chung@advantech.com.tw", "myadvantech@advantech.com", "BackOrder(" + dt.Rows(i)("SoldTo") + ") send failed:" + ex.Message.ToString(), "", False, "", "", Nothing, "")
            End Try

            'Frank 20160426 only send file one time when testing the program
            'Exit For

        Next

        Exit Function



        'ERPSQL.AppendFormat(" Select  STUFF ( ( ")
        'ERPSQL.AppendFormat(" SELECT distinct ','''+ COMPANY_ID  +'''' FROM   AEU_WEEKLY_BACKLOG_SCHEDULE  ")
        'ERPSQL.AppendFormat(" WHERE  WEEKDAY= Datepart(weekday, getdate() + @@DateFirst - 1)  ")
        'ERPSQL.AppendFormat(" for xml path('')  ")
        'ERPSQL.AppendFormat(" ),1,1,'') AS  ERPLIST ")


        Return dt
    End Function
    Private Function GetBackOrderDetail(ByVal companyID As String) As DataTable
        Dim sb As New StringBuilder()
        'sb.AppendFormat("   SELECT  BK.ORDERNO, BK.PONO, BK.SHIPTOID,")
        'sb.AppendFormat("   CONVERT(varchar(100), convert(datetime,BK.ORDERDATE,112), 101) as ORDERDATE , ")
        'sb.AppendFormat("  case BK.DELIVERYFLAG  when 'X' then 'Y' ELSE 'N' END as COMPLETE_DELIVERY_FLAG , BK.ORDERLINE,  ")
        'sb.AppendFormat("  BK.PRODUCTID,BK.ORDER_QTY as OrderQty, CONVERT(varchar(100),convert(datetime,BK.ORIGINALDD,112), 101) as REQUESTDATA, ")
        'sb.AppendFormat("  CONVERT(varchar(100),convert(datetime,BK.DUEDATE,112), 101) as DUEDATE,")
        'sb.AppendFormat("  ceiling(BK.SCHDLINECONFIRMQTY) as ConfirmedQTY, ceiling(BK.SCHDLINEOPENQTY) as OpenQTY,  ")
        'sb.AppendFormat("  BK.SCHEDLINESHIPEDQTY as ShippedQTY,BK.UNITPRICE, BK.TOTALPRICE,BK.CURRENCY, ")
        'sb.AppendFormat("  BK.DLV_QTY, BK.COMPANY_NAME ")
        'sb.AppendFormat("  FROM   (   ")
        'sb.AppendFormat("  select  ISNULL( C.PARENTCOMPANYID,AB.BILLTOID) AS SoldTo ,AB.*  ")
        'sb.AppendFormat("  from  SAP_BACKORDER_AB AB LEFT JOIN  SAP_DIMCOMPANY  C  ")
        'sb.AppendFormat("  ON AB.BILLTOID = C.COMPANY_ID  where C.ORG_ID ='EU10' and AB.SHIPTOID <> '' and AB.SHIPTOID is not null  ")
        'sb.AppendFormat("  ) AS BK INNER JOIN  ")
        'sb.AppendFormat("  SAP_DIMCOMPANY AS C ON BK.SHIPTOID = C.COMPANY_ID ")
        'sb.AppendFormat("  WHERE    (C.ORG_ID IN ('EU10')) AND (BK.SoldTo = '{0}')  and SCHDLINECONFIRMQTY > 0", companyID)
        'sb.AppendFormat("  ORDER BY   BK.ORDERNO, BK.ORDERLINE,BK.ORIGINALDD ")

        sb.AppendFormat("   SELECT  BK.ORDERNO, BK.PONO, BK.SHIPTOID,")
        sb.AppendFormat("   CONVERT(varchar(100), convert(datetime,BK.ORDERDATE,112), 101) as ORDERDATE , ")
        sb.AppendFormat("  case BK.DELIVERYFLAG  when 'X' then 'Y' ELSE 'N' END as COMPLETE_DELIVERY_FLAG , BK.ORDERLINE,  ")
        'sb.AppendFormat("  BK.PRODUCTID,BK.ORDER_QTY as OrderQty, CONVERT(varchar(100),convert(datetime,BK.ORIGINALDD,112), 101) as REQUESTDATA, ")
        sb.AppendFormat("  BK.PRODUCTID, BK.CUSTOMERPN, BK.ORDER_QTY as OrderQty,  ")


        sb.AppendFormat("  CASE WHEN BK.ROW=1 AND BK.SCHDLINECONFIRMQTY=0  ")
        sb.AppendFormat("  THEN ''  ")
        sb.AppendFormat("  ELSE CONVERT(varchar(100),convert(datetime,BK.ORIGINALDD,112), 101) ")
        sb.AppendFormat("  END AS REQUESTDATA, ")

        sb.AppendFormat("  CONVERT(varchar(100),convert(datetime,BK.DUEDATE,112), 101) as DUEDATE,")
        'sb.AppendFormat("  ceiling(BK.SCHDLINECONFIRMQTY) as ConfirmedQTY, ceiling(BK.SCHDLINEOPENQTY) as OpenQTY,  ")
        sb.AppendFormat("  ceiling(BK.SCHDLINECONFIRMQTY) as ConfirmedQTY,   ")

        sb.AppendFormat("  CASE WHEN BK.ORDER_QTY>0 AND BK.SCHEDLINESHIPEDQTY>0 THEN BK.ORDER_QTY-BK.SCHEDLINESHIPEDQTY   ")
        sb.AppendFormat("  ELSE 0   ")
        sb.AppendFormat("  END AS OpenQTY,   ")

        sb.AppendFormat("  BK.SCHEDLINESHIPEDQTY as ShippedQTY,BK.UNITPRICE, BK.TOTALPRICE,BK.CURRENCY, ")
        sb.AppendFormat("  BK.DLV_QTY, BK.COMPANY_NAME ")
        sb.AppendFormat("  FROM   (   ")
        sb.AppendFormat("  select  ISNULL( C.PARENTCOMPANYID,AB.BILLTOID) AS SoldTo ,AB.*  ")
        sb.AppendFormat("  ,ROW_NUMBER() OVER (partition by AB.OrderNo, AB.orderline ORDER BY AB.schdlineconfirmQTY DESC) AS ROW   ")
        sb.AppendFormat("  from  SAP_BACKORDER_AB AB LEFT JOIN  SAP_DIMCOMPANY  C  ")
        sb.AppendFormat("  ON AB.BILLTOID = C.COMPANY_ID  where C.ORG_ID ='EU10' and AB.SHIPTOID <> '' and AB.SHIPTOID is not null  ")
        sb.AppendFormat("  ) AS BK INNER JOIN  ")
        sb.AppendFormat("  SAP_DIMCOMPANY AS C ON BK.SHIPTOID = C.COMPANY_ID ")
        'sb.AppendFormat("  WHERE    (C.ORG_ID IN ('EU10')) AND (BK.SoldTo = '{0}')  and SCHDLINECONFIRMQTY > 0", companyID)
        sb.AppendFormat("  WHERE    (C.ORG_ID IN ('EU10')) AND (BK.SoldTo = '{0}')  ", companyID)
        sb.AppendFormat("  and (  ")
        sb.AppendFormat("  ( BK.ROW=1 AND BK.SCHDLINECONFIRMQTY=0 )   ")
        sb.AppendFormat("    OR  ")
        sb.AppendFormat("    BK.SCHDLINECONFIRMQTY>0  ")
        sb.AppendFormat("  )  ")
        sb.AppendFormat("  ORDER BY   BK.ORDERNO, BK.ORDERLINE,BK.ORIGINALDD ")


        Dim dt As DataTable = dbUtil.dbGetDataTable(CenterLibrary.DBConnection.MyAdvantechGlobal, sb.ToString())
        Return dt
    End Function
    Private Function GetSalesMails(ByVal companyID As String) As String
        Dim sb As New StringBuilder()
        sb.AppendFormat("   SELECT distinct ';'+ EMAIL  FROM ( ")
        sb.AppendFormat("   select DISTINCT   E.EMAIL  ")
        sb.AppendFormat("   from  dbo.SAP_COMPANY_EMPLOYEE CE INNER JOIN SAP_EMPLOYEE E ")
        sb.AppendFormat("   ON E.SALES_CODE = CE.SALES_CODE WHERE  ")
        sb.AppendFormat("   (CE.COMPANY_ID='{0}'   ) ", companyID)
        sb.AppendFormat("   AND  CE.SALES_ORG IN ('EU10')   ")
        sb.AppendFormat("   AND dbo.IsEmail(E.EMAIL)=1 and CE.PARTNER_FUNCTION in ('Z2','ZM') ")
        sb.AppendFormat("   AND E.EMAIL like '%@advantech.%' ")
        sb.AppendFormat("   UNION  ")
        sb.AppendFormat("   select  DISTINCT CONTACT_EMAIL AS EMAIL  from   dbo.SAP_COMPANY_CONTACTS  WHERE COMPANY_ID ='{0}' ", companyID)
        sb.AppendFormat("   AND  dbo.IsEmail(CONTACT_EMAIL)=1  ")
        sb.AppendFormat("   AND CONTACT_EMAIL like '%@advantech.%' ")
        sb.AppendFormat("   )  AS t ")
        sb.AppendFormat("   for xml path('') ")
        Dim mails As Object = dbUtil.dbExecuteScalar(CenterLibrary.DBConnection.MyAdvantechGlobal, sb.ToString())
        If mails IsNot Nothing AndAlso Not String.IsNullOrEmpty(mails) AndAlso mails.ToString.Trim <> ";" Then
            Return mails.ToString.Trim.Trim(";")
        End If
        Return ""
    End Function


    'Private Function GetBackOrder() As DataTable

    '    Dim CompanyList As DataTable = dbUtil.dbGetDataTable("MYLOCAL", "    select COMPANY_ID as CompanyID,Tomails,IsCallSAP   from  AEU_WEEKLY_BACKLOG_SCHEDULE")
    '    Dim ERPSQL As New StringBuilder()
    '    ERPSQL.AppendFormat(" Select  STUFF ( ( ")
    '    ERPSQL.AppendFormat(" SELECT distinct ','''+ COMPANY_ID  +'''' FROM   AEU_WEEKLY_BACKLOG_SCHEDULE  ")
    '    ERPSQL.AppendFormat(" WHERE  WEEKDAY= Datepart(weekday, getdate() + @@DateFirst - 1)  ")
    '    ERPSQL.AppendFormat(" for xml path('')  ")
    '    ERPSQL.AppendFormat(" ),1,1,'') AS  ERPLIST ")
    '    Dim Erpids As Object = dbUtil.dbExecuteScalar("MYLOCAL", ERPSQL.ToString())
    '    If Erpids Is Nothing OrElse String.IsNullOrEmpty(Erpids) Then
    '        dbUtil.SendEmailWithAttachment("ming.zhao@advantech.com.cn", "myadvantech@advantech.com", "BackOrder(Mails) send :  has no corresponding CompanyIDs ", "", False, "", "", Nothing, "")
    '        Exit Function
    '    End If
    '    Dim sb As New StringBuilder()
    '    sb.AppendFormat("   select distinct  T.SoldTo ,T.BILLTOID,T.COMPANYNAME FROM  ")
    '    sb.AppendFormat("   ( select  ISNULL( C.PARENTCOMPANYID,AB.BILLTOID) AS SoldTo ,AB.BILLTOID,C.COMPANY_NAME as COMPANYNAME ")
    '    sb.AppendFormat("    from  SAP_BACKORDER_AB AB LEFT JOIN  SAP_DIMCOMPANY  C  ")
    '    sb.AppendFormat("  ON AB.BILLTOID = C.COMPANY_ID  where C.ORG_ID ='EU10' and AB.SHIPTOID <> '' and AB.SHIPTOID is not null  ")
    '    sb.AppendFormat("   ) AS T ")
    '    sb.AppendFormat("   WHERE  T.SoldTo IN ({0})  ", Erpids)
    '    Dim dt As DataTable = dbUtil.dbGetDataTable("MY", sb.ToString())
    '    Dim dtDetai As New DataTable
    '    For i As Integer = 0 To dt.Rows.Count - 1
    '        Try
    '            dtDetai.Clear()
    '            dtDetai = GetBackOrderDetail(dt.Rows(i)("SoldTo"))
    '            Console.Write(dt.Rows(i)("SoldTo"))
    '            Console.WriteLine()
    '            Dim exePath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
    '            Dim strFPath As String = exePath + String.Format("\{0}.xls", dt.Rows(i)("SoldTo"))
    '            If IO.File.Exists(strFPath) = True Then
    '                IO.File.Delete(strFPath)
    '            End If
    '            dbUtil.DataTable2ExcelFile(dtDetai, strFPath)
    '            Dim fileStream As New FileStream(strFPath, FileMode.Open, FileAccess.Read, FileShare.Read)
    '            Dim bytes As Byte() = New Byte(fileStream.Length - 1) {}
    '            fileStream.Read(bytes, 0, bytes.Length)
    '            fileStream.Close()
    '            Dim stream As Stream = New MemoryStream(bytes)
    '            Dim Subject As String = String.Format("BackOrder for {0}({1})", dt.Rows(i)("COMPANYNAME"), dt.Rows(i)("SoldTo"))
    '            Dim dr() As DataRow = CompanyList.Select(String.Format("CompanyID='{0}'", dt.Rows(i)("SoldTo")))
    '            Dim salesmails As String = ""
    '            If dr.Length > 0 Then
    '                If Not IsDBNull(dr(0).Item("Tomails")) Then
    '                    salesmails = dr(0).Item("Tomails").ToString().Trim().Trim(New Char() {";"})
    '                End If
    '                If Not IsDBNull(dr(0).Item("IsCallSAP")) AndAlso dr(0).Item("IsCallSAP") = "1" Then
    '                    salesmails = salesmails + ";" + GetSalesMails(dt.Rows(i)("SoldTo"))
    '                End If
    '            End If
    '            salesmails = salesmails.ToString().Trim().Trim(New Char() {";"})
    '            Dim body As String = String.Format("This is a test for IT, it will be sent to "" {0} """, salesmails)
    '            dbUtil.SendEmailWithAttachment("Louis.Lin@advantech.eu;erika.molnarova@advantech.nl;myadvantech@advantech.com", "myadvantech@advantech.com", Subject, body, False, "", "", stream, String.Format("{0}.xls", dt.Rows(i)("SoldTo")))
    '        Catch ex As Exception
    '            Console.Write(ex.Message.ToString)
    '            Console.WriteLine()
    '            dbUtil.SendEmailWithAttachment("ming.zhao@advantech.com.cn", "myadvantech@advantech.com", "BackOrder(" + dt.Rows(i)("SoldTo") + ") send failed:" + ex.Message.ToString(), "", False, "", "", Nothing, "")
    '        End Try
    '    Next
    '    Return dt
    'End Function
    'Private Function GetBackOrderDetail(ByVal companyID As String) As DataTable
    '    Dim sb As New StringBuilder()
    '    sb.AppendFormat("   SELECT  BK.ORDERNO, BK.PONO, BK.SHIPTOID,")
    '    sb.AppendFormat("   CONVERT(varchar(100), convert(datetime,BK.ORDERDATE,112), 101) as ORDERDATE , ")
    '    sb.AppendFormat("  BK.CURRENCY, BK.ORDERLINE,  ")
    '    sb.AppendFormat("  BK.PRODUCTID, ceiling(BK.SCHDLINECONFIRMQTY) as SCHDLINECONFIRMQTY, ")
    '    sb.AppendFormat("  ceiling(BK.SCHDLINEOPENQTY) as SCHDLINEOPENQTY,  ")
    '    sb.AppendFormat("  BK.UNITPRICE, BK.TOTALPRICE,  ")
    '    sb.AppendFormat("  CONVERT(varchar(100), convert(datetime,BK.DUEDATE,112), 101) as DUEDATE , ")
    '    sb.AppendFormat("  BK.SCHEDLINESHIPEDQTY, BK.DLV_QTY, BK.COMPANY_NAME ")
    '    sb.AppendFormat("  FROM   (   ")
    '    sb.AppendFormat("  select  ISNULL( C.PARENTCOMPANYID,AB.BILLTOID) AS SoldTo ,AB.*  ")
    '    sb.AppendFormat("  from  SAP_BACKORDER_AB AB LEFT JOIN  SAP_DIMCOMPANY  C  ")
    '    sb.AppendFormat("  ON AB.BILLTOID = C.COMPANY_ID  where C.ORG_ID ='EU10' and AB.SHIPTOID <> '' and AB.SHIPTOID is not null  ")
    '    sb.AppendFormat("  ) AS BK INNER JOIN  ")
    '    sb.AppendFormat("  SAP_DIMCOMPANY AS C ON BK.SHIPTOID = C.COMPANY_ID ")
    '    sb.AppendFormat("  WHERE    (C.ORG_ID IN ('EU10')) AND (BK.SoldTo = '{0}') ", companyID)
    '    sb.AppendFormat("  ORDER BY   BK.ORDERNO, BK.ORDERLINE ")
    '    Dim dt As DataTable = dbUtil.dbGetDataTable("MY", sb.ToString())
    '    Return dt
    'End Function
    'Private Function GetSalesMails(ByVal companyID As String) As String
    '    Dim sb As New StringBuilder()
    '    sb.AppendFormat("   SELECT distinct ';'+ EMAIL  FROM ( ")
    '    sb.AppendFormat("   select DISTINCT   E.EMAIL  ")
    '    sb.AppendFormat("   from  dbo.SAP_COMPANY_EMPLOYEE CE INNER JOIN SAP_EMPLOYEE E ")
    '    sb.AppendFormat("   ON E.SALES_CODE = CE.SALES_CODE WHERE  ")
    '    sb.AppendFormat("   (CE.COMPANY_ID='{0}'   ) ", companyID)
    '    sb.AppendFormat("   AND  CE.SALES_ORG IN ('EU10')   ")
    '    sb.AppendFormat("   AND dbo.IsEmail(E.EMAIL)=1 and CE.PARTNER_FUNCTION in ('Z2','ZM') ")
    '    sb.AppendFormat("   AND E.EMAIL like '%@advantech.%' ")
    '    sb.AppendFormat("   UNION  ")
    '    sb.AppendFormat("   select  DISTINCT CONTACT_EMAIL AS EMAIL  from   dbo.SAP_COMPANY_CONTACTS  WHERE COMPANY_ID ='{0}' ", companyID)
    '    sb.AppendFormat("   AND  dbo.IsEmail(CONTACT_EMAIL)=1  ")
    '    sb.AppendFormat("   AND CONTACT_EMAIL like '%@advantech.%' ")
    '    sb.AppendFormat("   )  AS t ")
    '    sb.AppendFormat("   for xml path('') ")
    '    Dim mails As Object = dbUtil.dbExecuteScalar("MY", sb.ToString())
    '    If mails IsNot Nothing AndAlso Not String.IsNullOrEmpty(mails) AndAlso mails.ToString.Trim <> ";" Then
    '        Return mails.ToString.Trim.Trim(";")
    '    End If
    '    Return ""
    'End Function
End Module
