Imports System.Data.SqlClient
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
            dbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync CBOM From Estore have a error. " + Now.ToString("yyyy-MM-dd HH:mm:ss"), strSqlCmd, False)
            g_adoConn.Close() : Throw New Exception(ex.ToString + " sql:" + strSqlCmd)
        End Try
        'End Using
        g_adoConn.Close() : Return retInt
    End Function
    Public Shared Sub SendEmail( _
          ByVal SendTo As String, ByVal From As String, _
          ByVal Subject As String, ByVal Body As String, _
          ByVal IsBodyHtml As Boolean)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml

        mySmtpClient = New Net.Mail.SmtpClient(CenterLibrary.AppConfig.SMTPServerIP)
        Try
            mySmtpClient.Send(htmlMessage)
        Catch ex As System.Net.Mail.SmtpException
            System.Threading.Thread.Sleep(100)
            mySmtpClient.Send(htmlMessage)
        End Try
    End Sub
End Class
