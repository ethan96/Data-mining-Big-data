Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Text

Module Module1
    Dim InvalidOrgs As String = "'CN02','CN11','CN12','CN13','CN20','EU20','EU31','EU33','EU34','TW02','TWCP'"
    Sub Main(ByVal args() As String)
        Dim sm As New Net.Mail.SmtpClient(CenterLibrary.AppConfig.SMTPServerIP)
        Dim RunOrgs() As String = Nothing
        If args IsNot Nothing AndAlso args.Length = 1 Then
            RunOrgs = Split(args(0), ",")
        Else
            RunOrgs = New String() {"EU"}
        End If
        For Each RunOrg As String In RunOrgs
            Dim strOrgPrefix As String = RunOrg
            Console.WriteLine("Refreshing ATP of Org " + strOrgPrefix)
            Dim ws As New EQWS.quoteExit
            'Console.Read()
            Try
                Dim pdt As DataTable = GetStatusOPNTable(strOrgPrefix)
                Dim resultDt As New DataTable
                With resultDt.Columns
                    .Add("PART_NO") : .Add("SALES_ORG") : .Add("DLV_PLANT") : .Add("ATP_QTY", GetType(Integer))
                End With
                Dim p1 As New GET_MATERIAL_ATP.GET_MATERIAL_ATP
                p1.Connection = New SAP.Connector.SAPConnection(CenterLibrary.AppConfig.SAP_PRD)
                p1.Connection.Open()
                For Each pr As DataRow In pdt.Rows
                    Dim pn As String = pr.Item("PART_NO"), plant As String = pr.Item("DLV_PLANT")
                    Dim intATP As Integer = GetATP(pn, plant, p1)
                    If intATP > 0 Then
                        If Not IsPTD(pn) Then
                            Console.WriteLine("pn:" + pn + " qty:" + intATP.ToString())
                            Dim nr As DataRow = resultDt.NewRow()
                            nr.Item("PART_NO") = pn : nr.Item("SALES_ORG") = pr.Item("SALES_ORG")
                            nr.Item("DLV_PLANT") = plant : nr.Item("ATP_QTY") = intATP
                            resultDt.Rows.Add(nr)
                        End If
                    End If
                Next

                Dim conn As New SqlConnection(CenterLibrary.DBConnection.MyAdvantechGlobal)
                Dim bk As New SqlBulkCopy(conn)
                bk.DestinationTableName = "SAP_PRODUCT_ATP"
                Dim cmd As New SqlCommand("delete from SAP_PRODUCT_ATP where left(sales_org,2)='" + strOrgPrefix + "'", conn)
                conn.Open() : cmd.ExecuteNonQuery()
                bk.WriteToServer(resultDt)

                Dim dtOrderable As New DataTable
                'Dim strOrderablePN As String = _
                '    " SELECT a.PART_NO, a.SALES_ORG, a.DIST_CHANNEL, a.PRODUCT_STATUS, a.MIN_ORDER_QTY, a.MIN_DLV_QTY, a.MIN_BTO_QTY, a.DLV_PLANT, " + _
                '    " a.MATERIAL_PRICING_GRP, a.VALID_DATE, a.ITEM_CATEGORY_GROUP " + _
                '    " FROM SAP_PRODUCT_STATUS AS a left JOIN SAP_PRODUCT_ATP AS b ON a.PART_NO = b.PART_NO " + _
                '    " AND a.DLV_PLANT = b.DLV_PLANT AND a.SALES_ORG = b.SALES_ORG " + _
                '    " WHERE (((a.PRODUCT_STATUS = 'O') AND (b.ATP_QTY IS NOT NULL) AND (b.ATP_QTY > 0)) or a.PRODUCT_STATUS in ('A','N','H')) " + _
                '    " and left(a.sales_org,2)='" + strOrgPrefix + "' and a.sales_org not in (" + InvalidOrgs + ") "

                Dim strOrderablePN As New StringBuilder
                strOrderablePN.Append(" SELECT a.PART_NO, a.SALES_ORG, a.DIST_CHANNEL ")
                'Frank 2012/10/05
                Select Case strOrgPrefix.ToUpper
                    Case "US"
                        'For US, If part status is blank then column "proudct_status" will be written by A
                        strOrderablePN.Append(" , case when len(isnull(PRODUCT_STATUS,''))=0 then 'A' else PRODUCT_STATUS end as PRODUCT_STATUS ")
                    Case Else
                        strOrderablePN.Append(" , PRODUCT_STATUS ")
                End Select
                strOrderablePN.Append(" , a.MIN_ORDER_QTY, a.MIN_DLV_QTY, a.MIN_BTO_QTY, a.DLV_PLANT ")
                strOrderablePN.Append(" , a.MATERIAL_PRICING_GRP, a.VALID_DATE, a.ITEM_CATEGORY_GROUP ")
                strOrderablePN.Append(" FROM SAP_PRODUCT_STATUS AS a left JOIN SAP_PRODUCT_ATP AS b ON a.PART_NO = b.PART_NO ")
                strOrderablePN.Append(" AND a.DLV_PLANT = b.DLV_PLANT AND a.SALES_ORG = b.SALES_ORG ")
                strOrderablePN.Append(" WHERE ")
                'Frank 2012/10/05
                Select Case strOrgPrefix.ToUpper
                    Case "EU"
                        'IC & Ryan 20180426 EU80 status O parts but with avaialable stock should be able to order
                        'For EU, parts will be excluded if status is phase out.
                        strOrderablePN.Append(" (((a.SALES_ORG = 'EU80') AND (a.PRODUCT_STATUS = 'O') AND (b.ATP_QTY IS NOT NULL) AND (b.ATP_QTY > 0)) or a.PRODUCT_STATUS in ('A','N','H','M1','C','P','S2','S5','T','V')) ")
                    Case "US"
                        'For US, parts will be included if status is blank.
                        strOrderablePN.Append(" (((a.PRODUCT_STATUS = 'O') AND (b.ATP_QTY IS NOT NULL) AND (b.ATP_QTY > 0)) or a.PRODUCT_STATUS in ('','A','N','H','M1','C','P','S2','S5','T','V')) ")
                    Case Else
                        strOrderablePN.Append(" (((a.PRODUCT_STATUS = 'O') AND (b.ATP_QTY IS NOT NULL) AND (b.ATP_QTY > 0)) or a.PRODUCT_STATUS in ('A','N','H','M1','C','P','S2','S5','T','V')) ")
                End Select

                strOrderablePN.Append(" and left(a.sales_org,2)='" + strOrgPrefix + "' and a.sales_org not in (" + InvalidOrgs + ") ")

                'Dim apt As New SqlClient.SqlDataAdapter(strOrderablePN, conn)
                Dim apt As New SqlClient.SqlDataAdapter(strOrderablePN.ToString, conn)
                If conn.State <> ConnectionState.Open Then conn.Open()
                apt.Fill(dtOrderable)
                If dtOrderable.Rows.Count > 0 Then
                    cmd.CommandText = "delete from SAP_PRODUCT_STATUS_ORDERABLE where sales_org like '" + strOrgPrefix + "%' "
                    cmd.CommandTimeout = 5 * 60
                    If conn.State <> ConnectionState.Open Then conn.Open()
                    cmd.ExecuteNonQuery()
                    bk.DestinationTableName = "SAP_PRODUCT_STATUS_ORDERABLE"
                    If conn.State <> ConnectionState.Open Then conn.Open()
                    bk.WriteToServer(dtOrderable)
                End If

                'Ryan 20180416 Comment below out for ADLoG launching.
                'Frank 2012/05/14
                'Removing semi-finished products that under DLGR product line.
                'cmd.CommandText = "Delete from SAP_PRODUCT_STATUS_ORDERABLE where PART_NO in (Select PART_NO From sap_product where product_line='DLGR' and PRODUCT_TYPE<>'ZFIN') "
                'If conn.State <> ConnectionState.Open Then conn.Open()
                'cmd.ExecuteNonQuery()

                conn.Close()
                p1.Connection.Close()

            Catch ex As Exception
                'sm.Send("myadvantech@advantech.com", "tc.chen@advantech.com.tw,nada.liu@advantech.com.cn,frank.chung@advantech.com.tw", _
                '        "Sync Status O product ATP from SAP failed of " + strOrgPrefix, ex.ToString())
                sm.Send("MyAdvantech@advantech.com", "MyAdvantech@advantech.com", _
                        "Sync Orderable product ATP from SAP failed of " + strOrgPrefix, ex.ToString())
            End Try
        Next
        sm.Send("MyAdvantech@advantech.com", "MyAdvantech@advantech.com", _
                       "Sync Orderable product ATP from SAP successfully of " + String.Join(",", RunOrgs), "")
    End Sub

    Function GetATP(ByVal PartNo As String, ByVal Plant As String, ByRef p1 As GET_MATERIAL_ATP.GET_MATERIAL_ATP) As Integer
        Dim Inventory As Integer = 0
        PartNo = Format2SAPItem(Trim(UCase(PartNo)))
        Dim culQty As Integer = 0
        Dim retTb As New GET_MATERIAL_ATP.BAPIWMDVSTable, atpTb As New GET_MATERIAL_ATP.BAPIWMDVETable, rOfretTb As New GET_MATERIAL_ATP.BAPIWMDVS
        rOfretTb.Req_Qty = 9999 : rOfretTb.Req_Date = Now.ToString("yyyyMMdd")
        retTb.Add(rOfretTb)
        p1.Bapi_Material_Availability("", "A", "", New Short, "", "", "", PartNo, UCase(Plant), "", "", "", "", "PC", "", Inventory, "", "", _
                                      New GET_MATERIAL_ATP.BAPIRETURN, atpTb, retTb)
        Inventory = 0
        Dim atpDt As DataTable = atpTb.ToADODataTable()
        For Each r As DataRow In atpDt.Rows
            Inventory += CType(r.Item("com_qty"), Integer)
        Next
        Return Inventory
    End Function

    Function GetStatusOPNTable(ByVal OrgPrefix As String) As DataTable
        Dim strSql As String = _
            " select distinct top 99999 a.PART_NO, a.DLV_PLANT, a.SALES_ORG  " + _
            " from SAP_PRODUCT_STATUS a  " + _
            " where left(a.sales_org,'2')='" + OrgPrefix + "' and a.PRODUCT_STATUS ='O' and a.DLV_PLANT not in " + _
            " (' ','CDB6','CDM6','CKB1','CKB3','CKB4','CKB5','CKB6','CKB7','CKM3','CKM4','CKM6','EPH1') " + _
            " and a.DLV_PLANT is not null  " + _
            " and a.SALES_ORG not in (" + InvalidOrgs + ") and a.SALES_ORG is not null " + _
            " and LEFT(a.part_no,1) not in ('#','$')  " + _
            " order by a.PART_NO  "
        Dim conn As New SqlConnection(CenterLibrary.DBConnection.MyAdvantechGlobal)
        Dim apt As New SqlDataAdapter(strSql, conn), dt As New DataTable
        apt.Fill(dt)
        Return dt
    End Function

    Function IsPTD(ByVal PartNo As String) As Boolean
        Dim f As Boolean = False
        Dim STR As String = String.Format("select * from SAP_PRODUCT where " & _
                                            " ((PRODUCT_TYPE = 'ZPER') " & _
                                            " OR " & _
                                            " ((PRODUCT_TYPE = 'ZFIN' OR PRODUCT_TYPE = 'ZOEM') AND (PART_NO LIKE 'BT%' OR PART_NO LIKE 'DSD%' OR PART_NO LIKE 'ES%' OR PART_NO LIKE 'EWM%' OR PART_NO LIKE 'GPS%' OR PART_NO LIKE 'SQF%' OR PART_NO LIKE 'WIFI%' OR PART_NO LIKE 'PMM%' OR PART_NO LIKE 'Y%')) " & _
                                            " OR " & _
                                            " ((PRODUCT_TYPE = 'ZRAW') AND (PART_NO LIKE '206Q%')) " & _
                                            " OR " & _
                                            " ((PRODUCT_TYPE = 'ZSEM') AND (PART_NO LIKE '968Q%'))) AND PART_NO = '{0}'", PartNo)
        Dim DT As New DataTable
        Dim conn As New SqlConnection(CenterLibrary.DBConnection.MyAdvantechGlobal)
        Dim apt As New SqlDataAdapter(STR, conn)
        apt.Fill(DT)
        If DT.Rows.Count > 0 Then
            f = True
        End If
        Return f
    End Function

    Public Function Format2SAPItem(ByVal Part_No As String) As String

        Try
            If IsNumericItem(Part_No) And Not Part_No.Substring(0, 1).Equals("0") Then
                Dim zeroLength As Integer = 18 - Part_No.Length
                For i As Integer = 0 To zeroLength - 1
                    Part_No = "0" & Part_No
                Next
                Return Part_No
            Else
                Return Part_No
            End If
        Catch ex As Exception
            Return Part_No
        End Try

    End Function

    Public Function IsNumericItem(ByVal part_no As String) As Boolean

        Dim pChar() As Char = part_no.ToCharArray()

        For i As Integer = 0 To pChar.Length - 1
            If Not IsNumeric(pChar(i)) Then
                Return False
                Exit Function
            End If
        Next

        Return True
    End Function

End Module
