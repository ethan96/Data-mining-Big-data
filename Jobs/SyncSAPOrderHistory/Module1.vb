Imports System.Data.SqlClient

Module Module1
    Public intDiffDays As Integer = 7
    Sub Main()

        Dim mySmtpClient As New Net.Mail.SmtpClient(CenterLibrary.AppConfig.SMTPServerIP)
        mySmtpClient.Send("myadvantech@advantech.com", "myadvantech@advantech.com", "Start Sync SAP_ORDER_HISTORY", Now.ToString("yyyyMMddHHmmss"))
        Try
            Dim strSAP As String = CenterLibrary.DBConnection.SAP_PRD
            Dim strRFM As String = CenterLibrary.DBConnection.MyAdvantechGlobal
            Dim sb As New System.Text.StringBuilder
            With sb
                .AppendLine(" select a.KUNNR as COMPANY_ID, a.VBELN as SO_NO, a.VKORG as SALES_ORG, to_date(a.ERDAT,'yyyy/MM/dd') as ORDER_DATE,  ")
                .AppendLine(" b.POSNR as LINE_NO, b.UEPOS as HIGHER_LEVEL, b.MATNR as PART_NO, cast(b.KWMENG as integer) as ORDER_QTY, b.WAERK as CURRENCY, b.NETWR as UNIT_PRICE, " + _
                            " null as US_AMT, 0 as PRICE_FACTOR_UPDATED ")
                .AppendLine(" FROM SAPRDP.VBAK a INNER JOIN SAPRDP.VBAP b ON a.VBELN = b.VBELN ")
                .AppendLine(" where a.MANDT='168' and b.MANDT='168' and a.AUART in ('ZOR','ZOR2')  ")
                .AppendLine(" and (a.AEDAT>=to_char(sysdate-" + intDiffDays.ToString() + ",'yyyymmdd') or a.ERDAT>=to_char(sysdate-" + intDiffDays.ToString() + ",'yyyymmdd')) ")
            End With
            Dim dt As DataTable = dbGetDataTableOra(strSAP, sb.ToString())

            Dim sonoAry As New ArrayList
            For Each r As DataRow In dt.Rows

                r.Item("SO_NO") = RemoveZeroString(r.Item("SO_NO"))

                Dim strTmpSONO As String = "'" + r.Item("SO_NO").ToString() + "'"
                If Not sonoAry.Contains(strTmpSONO) Then sonoAry.Add(strTmpSONO)
            Next
            If sonoAry.Count > 0 Then
                Dim strSOInString As String = String.Join(",", sonoAry.ToArray())
                'MsgBox(strSOInString)
                Dim myConn As New SqlClient.SqlConnection(strRFM)
                myConn.Open()
                Dim cmd As New SqlClient.SqlCommand("delete from sap_order_history where so_no in (" + strSOInString + ") ", myConn)
                cmd.ExecuteNonQuery()
                Dim bk As New SqlBulkCopy(myConn)
                bk.DestinationTableName = "sap_order_history"
                bk.BulkCopyTimeout = 60 * 30
                If myConn.State <> ConnectionState.Open Then myConn.Open()
                bk.WriteToServer(dt)

                Dim strUpdateSql As String = _
                    " update SAP_ORDER_HISTORY set UNIT_PRICE=UNIT_PRICE * power(10,2-b.factor), PRICE_FACTOR_UPDATED=1  " + _
                    " from SAP_ORDER_HISTORY a inner join SAP_TCURX b on a.currency=b.currency  " + _
                    " where a.currency not in ('EUR','USD','GBP') and PRICE_FACTOR_UPDATED=0 " + _
                    " " + _
                    " update SAP_ORDER_HISTORY set PRICE_FACTOR_UPDATED=1  " + _
                    " from SAP_ORDER_HISTORY a inner join SAP_TCURX b on a.currency=b.currency  " + _
                    " where a.currency in ('EUR','USD','GBP') and PRICE_FACTOR_UPDATED=0 " + _
                    " " + _
                    " update SAP_ORDER_HISTORY set US_AMT = UNIT_PRICE * ORDER_QTY where CURRENCY='USD' and US_AMT is null  " + _
                    "  " + _
                    " update SAP_ORDER_HISTORY set US_AMT=UNIT_PRICE*ORDER_QTY* " + _
                    " IsNull((select top 1 UKURS from SAP_EXCHANGERATE b where b.fcurr=a.CURRENCY and b.tcurr='USD' and EXCH_DATE<=a.ORDER_DATE order by EXCH_DATE desc),0) " + _
                    " from SAP_ORDER_HISTORY a where a.CURRENCY<>'USD' and a.US_AMT is null  " + _
                    "  " + _
                    " update SAP_ORDER_HISTORY set PART_NO=dbo.DelPrevZero(PART_NO) where PART_NO like '0%' " + _
                    "  " + _
                    " update SAP_ORDER_HISTORY set LINE_NO=dbo.DelPrevZero(LINE_NO) where LINE_NO like '0%' " + _
                    "  " + _
                    " update SAP_ORDER_HISTORY set HIGHER_LEVEL=dbo.DelPrevZero(HIGHER_LEVEL) where HIGHER_LEVEL like '0%' "
                If myConn.State <> ConnectionState.Open Then myConn.Open()
                cmd = New SqlClient.SqlCommand(strUpdateSql, myConn)
                cmd.CommandTimeout = 99999
                cmd.ExecuteNonQuery()
                myConn.Close()
            End If
            mySmtpClient.Send("myadvantech@advantech.com", "myadvantech@advantech.com", "OK Sync SAP_ORDER_HISTORY", Now.ToString("yyyyMMddHHmmss"))

        Catch ex As Exception
            mySmtpClient.Send("myadvantech@advantech.com", "myadvantech@advantech.com", "Error Sync SAP_ORDER_HISTORY", ex.ToString())

        End Try
    End Sub

    Public Function dbGetDataTableOra(ByVal ConnectionName As String, ByVal strSqlCmd As String) As DataTable
        Dim g_adoConn As New Oracle.DataAccess.Client.OracleConnection(ConnectionName)

        Dim dt As New DataTable
        Dim da As New Oracle.DataAccess.Client.OracleDataAdapter(strSqlCmd, g_adoConn)

        Try
            da.Fill(dt)
        Catch ex As Exception
            g_adoConn.Close() : g_adoConn.Dispose()
            Threading.Thread.CurrentThread.Priority = Threading.ThreadPriority.Normal
            Throw ex
        End Try
        g_adoConn.Close() : g_adoConn = Nothing
        Return dt
    End Function

    Public Function RemoveZeroString(ByVal NumericPart_No As String) As String

        If IsNumericItem(NumericPart_No) Then
            For i As Integer = 0 To NumericPart_No.Length - 1
                If Not NumericPart_No.Substring(i, 1).Equals("0") Then
                    Return NumericPart_No.Substring(i)
                    Exit For
                End If
            Next
            Return NumericPart_No
        Else
            Return NumericPart_No
        End If

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
