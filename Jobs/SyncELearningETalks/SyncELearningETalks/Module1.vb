Module Module1
    Public strConnMy As String = CenterLibrary.DBConnection.MyAdvantechGlobal
    Public strConnElearning As String = CenterLibrary.DBConnection.Elearning

    Sub Main()
        Dim connMy As New SqlClient.SqlConnection(strConnMy)
        Dim connELearning As New SqlClient.SqlConnection(strConnElearning)
        connMy.Open() : connELearning.Open()
        Dim dt As New DataTable
        Dim apt As New SqlClient.SqlDataAdapter(" SELECT [Subject],[ChineseName],[EnglishName],[CourseSyllabus],[ENCourseSyllabus],[Speaker],[WistaPath],[CreateCourseTime] FROM [eLearning_Advantech].[dbo].[V_ETALKS_COURSE]", connELearning)
        apt.SelectCommand.CommandTimeout = 600 * 1000
        Try
            apt.Fill(dt)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim cmd As New SqlClient.SqlCommand("truncate table ELEARNING_ETALKS", connMy)
                cmd.ExecuteNonQuery()

                Dim bk As New SqlClient.SqlBulkCopy(connMy)
                bk.DestinationTableName = "ELEARNING_ETALKS"
                bk.WriteToServer(dt)
            End If

        Catch ex As Exception
            CenterLibrary.MailUtil.SendEmail("MyAdvantech@advantech.com", "MyAdvantech@advantech.com", "Sync eLearning eTalks Failed", ex.ToString, True, "", "")
        End Try
        connMy.Close() : connELearning.Close()
    End Sub

    Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal Subject As String, ByVal Body As String, ByVal IsBodyHtml As Boolean,
           ByVal cc As String, ByVal bcc As String, Optional ByVal NotifyOnFailure As Boolean = False)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml
        If cc <> "" Then htmlMessage.CC.Add(cc)
        Try
            If bcc <> "" Then htmlMessage.Bcc.Add(bcc)
        Catch ex As Exception
            Throw New Exception("BCC:" + bcc + " caused error for sending email")
        End Try

        If NotifyOnFailure Then htmlMessage.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.OnFailure
        'htmlMessage.CC.Add("tc.chen@advantech.com.tw")
        'htmlMessage.CC.Add("jackie.wu@advantech.com.cn")
        mySmtpClient = New Net.Mail.SmtpClient("172.20.0.76")
        mySmtpClient.Send(htmlMessage)
    End Sub
End Module
