Imports Microsoft.VisualBasic
Imports Oracle.DataAccess.Client
Public Class OraDbUtill
    Public Shared Function dbGetDataTable( _
ByVal ConnectionName As String, _
ByVal strSqlCmd As String) As DataTable

        Dim g_adoConn As New OracleConnection(ConnectionName)
        Dim dt As New DataTable
        Dim da As New OracleDataAdapter(strSqlCmd, g_adoConn)
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
