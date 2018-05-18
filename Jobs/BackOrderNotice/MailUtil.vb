Public Class MailUtil
    Public Shared Sub SendEmail( _
     ByVal SendTo As String, ByVal From As String, _
     ByVal Subject As String, ByVal Body As String, _
     ByVal IsBodyHtml As Boolean)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml
        'mySmtpClient = New Net.Mail.SmtpClient("172.21.34.21")
        mySmtpClient = New Net.Mail.SmtpClient("172.20.0.76")
        Try
            mySmtpClient.Send(htmlMessage)
        Catch ex As System.Net.Mail.SmtpException
            System.Threading.Thread.Sleep(100)
            mySmtpClient.Send(htmlMessage)
        End Try
    End Sub
End Class
