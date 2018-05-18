Imports System.Data.SqlClient
Imports System.IO
Imports System.Reflection
Imports Oracle.DataAccess.Client
Imports System.Configuration

Public Class dbUtil
    Public Shared Function dbOracleGetDataTable( _
ByVal ConnectionName As String, _
ByVal strSqlCmd As String, Optional CommandTimeoutInSecond As Integer = 0) As DataTable
        'Dim aa As New Oracle.DataAccess.Client.OracleConnection 
        Dim g_adoConn As New OracleConnection(ConfigurationManager.ConnectionStrings(ConnectionName).ConnectionString)
        Dim dt As New DataTable
        Dim da As New OracleDataAdapter(strSqlCmd, g_adoConn)
        'da.SelectCommand.CommandTimeout = 30
        If CommandTimeoutInSecond > 0 Then
            da.SelectCommand.CommandTimeout = CommandTimeoutInSecond
        End If
        Threading.Thread.CurrentThread.Priority = Threading.ThreadPriority.BelowNormal
        Try
            da.Fill(dt)
        Catch ex As Exception
            g_adoConn.Close() : g_adoConn.Dispose()
            Threading.Thread.CurrentThread.Priority = Threading.ThreadPriority.Normal
            Throw ex
        End Try
        g_adoConn.Close() : g_adoConn = Nothing
        Threading.Thread.CurrentThread.Priority = Threading.ThreadPriority.Normal
        Return dt
    End Function
    Public Shared Function dbGetDataTable( _
  ByVal ConnectionName As String, _
  ByVal strSqlCmd As String) As DataTable
        Dim g_adoConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(ConnectionName).ConnectionString)
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
        Dim g_adoConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(ConnectionName).ConnectionString)
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
        Dim g_adoConn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings(ConnectionStringName).ConnectionString)
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
        Dim oSmpt As New Net.Mail.SmtpClient("aeuht1.aeu.advantech.corp")
        ' Try
        oSmpt.Send(oMail)
        Return True
        ' Catch ex As Exception
        ' End Try
        Return False
    End Function
End Class
