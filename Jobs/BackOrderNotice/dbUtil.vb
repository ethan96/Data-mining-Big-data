Imports System.Data.SqlClient
Imports System.IO
Imports System.Reflection

Public Class dbUtil
    Public Shared Function dbGetDataTable( _
  ByVal ConnectionName As String, _
  ByVal strSqlCmd As String) As DataTable
        Dim g_adoConn As New SqlConnection(ConnectionName)
        Dim dt As New DataTable
        Dim da As New SqlDataAdapter(strSqlCmd, g_adoConn)
        da.SelectCommand.CommandTimeout = 5 * 60
        Try
            da.Fill(dt)
        Catch ex As Exception
            g_adoConn.Close() : Throw New Exception(ex.ToString() + vbTab + "sql:" + strSqlCmd)
        End Try
        g_adoConn.Close() : g_adoConn = Nothing
        Return dt
    End Function
    Public Shared Function dbExecuteScalar(ByVal ConnectionName As String, ByVal strSqlCmd As String) As Object
        Dim g_adoConn As New SqlConnection(ConnectionName)
        For i As Integer = 0 To 3
            Try
                g_adoConn.Open()
                Exit For
            Catch ex As SqlException
                If i = 3 Then Throw ex
                Threading.Thread.Sleep(100)
            End Try
        Next
        Dim dbCmd As SqlClient.SqlCommand = g_adoConn.CreateCommand()
        dbCmd.CommandType = CommandType.Text : dbCmd.CommandText = strSqlCmd : dbCmd.CommandTimeout = 5 * 60
        Dim retObj As Object = Nothing
        Try
            retObj = dbCmd.ExecuteScalar()
        Catch ex As Exception
            g_adoConn.Close() : Throw ex
        End Try
        g_adoConn.Close() : Return retObj
    End Function
    Public Shared Function dbExecuteNoQuery( _
   ByVal ConnectionStringName As String, _
   ByVal strSqlCmd As String) As Integer
        Dim g_adoConn As New SqlConnection(ConnectionStringName)
        Dim dbCmd As SqlClient.SqlCommand = g_adoConn.CreateCommand()
        dbCmd.Connection = g_adoConn : dbCmd.CommandText = strSqlCmd
        dbCmd.CommandTimeout = 10 * 60
        Dim retInt As Integer = -1
        For i As Integer = 0 To 3
            Try
                g_adoConn.Open()
                Exit For
            Catch ex As SqlException
                If i = 3 Then Throw ex
                Threading.Thread.Sleep(100)
            End Try
        Next
        'Using tran As SqlTransaction = g_adoConn.BeginTransaction
        Try
            'dbCmd.Transaction = tran
            retInt = dbCmd.ExecuteNonQuery()
            'tran.Commit()
        Catch ex As Exception
            'tran.Rollback()
            g_adoConn.Close() : Throw New Exception(ex.ToString + " sql:" + strSqlCmd)
        End Try
        'End Using
        g_adoConn.Close() : Return retInt
    End Function
    Public Shared Function DataTable2ExcelFile(ByVal dt As DataTable, ByVal path As String) As Boolean
        Dim license As Aspose.Cells.License = New Aspose.Cells.License()
        Dim exePath As String = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        Dim strFPath As String = exePath + "\Aspose.Total.lic"
        license.SetLicense(strFPath)
        'Try
        Dim wb As New Aspose.Cells.Workbook
        wb.Worksheets.Add(Aspose.Cells.SheetType.Worksheet)
        For i As Integer = 0 To dt.Columns.Count - 1
            wb.Worksheets(0).Cells(0, i).PutValue(dt.Columns(i).ColumnName)
        Next
        For i As Integer = 0 To dt.Rows.Count - 1
            For j As Integer = 0 To dt.Columns.Count - 1
                wb.Worksheets(0).Cells(i + 1, j).PutValue(dt.Rows(i).Item(j).ToString())
            Next
        Next
        wb.Save(path)
        'With HttpContext.Current.Response
        '    .Clear()
        '    .ContentType = "application/vnd.ms-excel"
        '    .AddHeader("Content-Disposition", String.Format("attachment; filename={0};", path))
        '    Try
        '        .BinaryWrite(wb.SaveToStream().ToArray)
        '    Catch ex As Exception
        '        .End()
        '    End Try
        'End With

        Return True
    End Function
    Public Shared Function DataTable2ExcelStream(ByVal dt As DataTable) As IO.MemoryStream
        Dim license As Aspose.Cells.License = New Aspose.Cells.License()
        Dim exePath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
        Dim strFPath As String = exePath + "\Aspose.Total.lic"
        license.SetLicense(strFPath)
        'Try
        Dim wb As New Aspose.Cells.Workbook
        wb.Worksheets.Add(Aspose.Cells.SheetType.Worksheet)
        For i As Integer = 0 To dt.Columns.Count - 1
            wb.Worksheets(0).Cells(0, i).PutValue(dt.Columns(i).ColumnName)
        Next
        For i As Integer = 0 To dt.Rows.Count - 1
            For j As Integer = 0 To dt.Columns.Count - 1
                If dt.Rows(i).Item(j).ToString.StartsWith("=") Then
                    wb.Worksheets(0).Cells(i + 1, j).Formula = dt.Rows(i).Item(j).ToString()
                Else
                    wb.Worksheets(0).Cells(i + 1, j).PutValue(dt.Rows(i).Item(j))
                End If
            Next
        Next
        Return wb.SaveToStream()
        'Catch ex As Exception
        '    Return Nothing
        'End Try
    End Function
    Public Shared Function SendEmailWithAttachment( _
            ByVal SendTo As String, ByVal From As String, _
            ByVal Subject As String, ByVal Body As String, _
            ByVal IsBodyHtml As Boolean, _
            ByVal cc As String, _
            ByVal bcc As String, ByVal AttachmentStreams As System.IO.Stream, ByVal AttachmentName As String) As Boolean
        Dim oMail As New Net.Mail.MailMessage()
        oMail.From = New Net.Mail.MailAddress("myadvantech@advantech.com")
        If SendTo.Contains(";") Then
            For Each emailadrr As String In SendTo.Split(";")
                oMail.To.Add(emailadrr.Trim())
            Next
        Else
            oMail.To.Add(SendTo.Trim())
        End If
        oMail.Bcc.Add("myadvantech@advantech.com")
            oMail.Subject = Subject
            oMail.IsBodyHtml = IsBodyHtml
        oMail.Body = Body
        If AttachmentStreams IsNot Nothing Then
            oMail.Attachments.Add(New Net.Mail.Attachment(AttachmentStreams, AttachmentName))
        End If
        Dim oSmpt As New Net.Mail.SmtpClient(CenterLibrary.AppConfig.SMTPServerIP)
        ' Try
        oSmpt.Send(oMail)
            Return True
            ' Catch ex As Exception
            ' End Try
            Return False
    End Function
End Class
