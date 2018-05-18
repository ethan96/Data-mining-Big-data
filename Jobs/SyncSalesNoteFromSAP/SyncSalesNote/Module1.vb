Imports Z_READ_TEXT
Imports System.Data.SqlClient
Module Module1
    Sub Main()

        Try
            'OraDbUtill.SendEmail("tc.chen@advantech.eu,ming.zhao@advantech.com.cn", "ebiz.aeu@advantech.eu", "Sync SalesNote From SAP start " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "", False)
            Dim strFMA As String = CenterLibrary.DBConnection.MyAdvantechGlobal
            Dim GlobalDT As New DataTable()
            GlobalDT.Columns.Add("COMPANY_ID", GetType(String))
            GlobalDT.Columns.Add("TXTTYPE", GetType(String))
            GlobalDT.Columns.Add("TDOBJECT", GetType(String))
            GlobalDT.Columns.Add("TXT", GetType(String))
            Dim OracleDt As DataTable = OraDbUtill.dbGetDataTable(CenterLibrary.DBConnection.SAP_PRD,
                "select TDNAME,TDID,TDOBJECT from SAPRDP.STXL where MANDT='168' AND RELID='TX' and (TDOBJECT ='KNA1' OR TDOBJECT ='MVKE' OR TDOBJECT ='KNVV') AND TDID='0001' and CLUSTD is not null and TDNAME IS NOT NULL AND TDSPRAS ='E'")

            'Frank test codes
            'Dim OracleDt As DataTable = OraDbUtill.dbGetDataTable("SAP_PRD", _
            '"select TDNAME,TDID,TDOBJECT from SAPRDP.STXL where MANDT='168' AND RELID='TX' and (TDOBJECT ='KNA1' OR TDOBJECT ='MVKE' OR TDOBJECT ='KNVV') AND TDID='0001' and CLUSTD is not null and TDNAME IS NOT NULL AND TDSPRAS ='E' and TDNAME='UWE46062'")


            If OracleDt.Rows.Count > 0 Then
                Dim eup As New Z_READ_TEXT.Z_READ_TEXT
                Dim pout1 As New Z_READ_TEXT.THEAD
                Dim pout2 As New Z_READ_TEXT.TLINETable
                eup.Connection = New SAP.Connector.SAPConnection(System.Configuration.ConfigurationManager.AppSettings("SAPConnPrd"))
                eup.Connection.Open()
                Dim m As Integer = 0
                For i As Integer = 0 To OracleDt.Rows.Count - 1
                    Dim dr As DataRow = GlobalDT.NewRow()
                    dr("COMPANY_ID") = OracleDt.Rows(i).Item("TDNAME").ToString.Trim.Replace("'", "''")
                    dr("TXTTYPE") = OracleDt.Rows(i).Item("TDID").ToString.Trim.Replace("'", "''")
                    dr("TDOBJECT") = OracleDt.Rows(i).Item("TDOBJECT").ToString.Trim.Replace("'", "''")
                    dr("TXT") = ""
                    Try
                        eup.Zread_Text(0, "168", "0001", "EN", "", OracleDt.Rows(i).Item("TDNAME").ToString.Trim.Replace("'", "''"), OracleDt.Rows(i).Item("TDOBJECT").ToString, pout1, pout2)
                    Catch ex As Exception
                        OraDbUtill.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SalesNote From SAP Error " + Now.ToString("yyyy-MM-dd HH:mm:ss"), OracleDt.Rows(i).Item("TDNAME").ToString + "<hr/>" + OracleDt.Rows(i).Item("TDOBJECT").ToString + "<hr/>" + ex.ToString(), False)
                        m += 1
                        If m = 20 Then Exit For
                    End Try
                    Dim ExportDT As DataTable = pout2.ToADODataTable()
                    If ExportDT.Rows.Count > 0 Then
                        For j As Integer = 0 To ExportDT.Rows.Count - 1
                            If Not IsDBNull(ExportDT.Rows(j).Item("Tdline")) Then
                                dr("TXT") += ExportDT.Rows(j).Item("Tdline").ToString.Trim.Replace("'", "''") + vbCrLf
                            End If
                        Next
                    End If
                    GlobalDT.Rows.Add(dr)
                Next
                eup.Connection.Close()
                GlobalDT.AcceptChanges()
                Dim salenoteDT As New DataTable, ordernoteDT As New DataTable
                salenoteDT = GlobalDT.Clone() : ordernoteDT = GlobalDT.Clone()
                Dim drs() As DataRow = GlobalDT.Select("TDOBJECT = 'KNA1'")
                For i As Integer = 0 To drs.Length - 1
                    salenoteDT.ImportRow(CType(drs(i), DataRow))
                Next
                salenoteDT.AcceptChanges()
                'salenoteDT.Columns.Remove("TDOBJECT")
                'salenoteDT.Columns.Add("LAST_UPD_DATE", GetType(Date))
                'For i As Integer = 0 To salenoteDT.Rows.Count - 1
                '    salenoteDT.Rows(i).Item("LAST_UPD_DATE") = Now()
                'Next
                'salenoteDT.AcceptChanges()
                'Ming add for knvv
                drs = GlobalDT.Select("TDOBJECT = 'KNVV'")
                For i As Integer = 0 To drs.Length - 1
                    salenoteDT.ImportRow(CType(drs(i), DataRow))
                Next
                salenoteDT.AcceptChanges()
                'salenoteDT.Columns.Remove("TDOBJECT")
                salenoteDT.Columns.Add("LAST_UPD_DATE", GetType(Date))
                For i As Integer = 0 To salenoteDT.Rows.Count - 1
                    If String.Equals(salenoteDT.Rows(i).Item("TDOBJECT"), "KNVV") Then
                        salenoteDT.Rows(i).Item("LAST_UPD_DATE") = Now.AddMinutes(5)
                    Else
                        salenoteDT.Rows(i).Item("LAST_UPD_DATE") = Now()
                    End If

                Next
                salenoteDT.Columns.Remove("TDOBJECT")
                salenoteDT.AcceptChanges()
                'end
                drs = GlobalDT.Select("TDOBJECT = 'MVKE'")
                For i As Integer = 0 To drs.Length - 1
                    ordernoteDT.ImportRow(CType(drs(i), DataRow))
                Next
                ordernoteDT.Columns.Remove("TDOBJECT")
                ordernoteDT.AcceptChanges()
                ordernoteDT.Columns.Add("LAST_UPD_DATE", GetType(Date))

                ordernoteDT.Columns.Add("ORG", GetType(String))
                ordernoteDT.Columns.Add("DISCHANNEL", GetType(String))

                For i As Integer = 0 To ordernoteDT.Rows.Count - 1
                    ordernoteDT.Rows(i).Item("LAST_UPD_DATE") = Now()
                    Dim TDNAME As String = ordernoteDT.Rows(i).Item("COMPANY_ID")
                    Dim ORGANDDIS As String = Right(ordernoteDT.Rows(i).Item("COMPANY_ID").ToString.Trim, 6)
                    Dim PARTNO As String = TDNAME.Replace(ORGANDDIS, "").Trim()
                    Dim ORG As String = Left(ORGANDDIS, 4)
                    Dim DIS As String = Right(ORGANDDIS, 2)

                    ordernoteDT.Rows(i).Item("ORG") = ORG
                    ordernoteDT.Rows(i).Item("DISCHANNEL") = DIS
                    ordernoteDT.Rows(i).Item("COMPANY_ID") = PARTNO.Trim.TrimStart("0")
                Next
                ordernoteDT.AcceptChanges()



                '--truncate table SAP_COMPANY_SALESNOTE
                Dim g_adoConn As New SqlConnection(strFMA)
                Dim dbCmd As SqlClient.SqlCommand = g_adoConn.CreateCommand()
                dbCmd.Connection = g_adoConn
                g_adoConn.Open()
                dbCmd.CommandText = "truncate table SAP_COMPANY_SALESNOTE;truncate table SAP_PRODUCT_ORDERNOTE"
                dbCmd.ExecuteNonQuery()
                g_adoConn.Close()
                '--end

                Dim sdt As New SqlClient.SqlBulkCopy(strFMA)
                sdt.DestinationTableName = "SAP_COMPANY_SALESNOTE"
                sdt.WriteToServer(salenoteDT)
                Dim odt As New SqlClient.SqlBulkCopy(strFMA)
                odt.DestinationTableName = "SAP_PRODUCT_ORDERNOTE"
                odt.WriteToServer(ordernoteDT)
            End If
        Catch ex As Exception
            OraDbUtill.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SalesNote From SAP Error " + Now.ToString("yyyy-MM-dd HH:mm:ss"), ex.ToString(), False)
            Exit Sub
        End Try
        OraDbUtill.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SalesNote From SAP OK " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "", False)

    End Sub

End Module
