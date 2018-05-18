Imports Microsoft.VisualBasic

Public Class MailUtil

    Public Shared Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal Subject As String, ByVal Body As String, ByVal IsBodyHtml As Boolean,
           ByVal cc As String, ByVal bcc As String)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml
        If cc <> "" Then htmlMessage.CC.Add(cc)
        Try
            If bcc <> "" Then htmlMessage.Bcc.Add(bcc)
        Catch ex As Exception
            Throw New Exception("BCC:" + bcc + " caused error for sending email")
        End Try

        'htmlMessage.CC.Add("tc.chen@advantech.com.tw")
        'htmlMessage.CC.Add("jackie.wu@advantech.com.cn")
        mySmtpClient = New Net.Mail.SmtpClient(AppConfig.SMTPServerIP)
        Try
            mySmtpClient.Send(htmlMessage)
        Catch ex As System.Net.Mail.SmtpException
            Dim wr As New WriteLog With {.FunctionName = Subject, .Message = Body}
            wr.WriteErrorLog()
        End Try
    End Sub


End Class

