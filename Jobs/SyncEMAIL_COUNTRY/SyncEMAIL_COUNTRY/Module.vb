Module Module1
    Public strMY As String = CenterLibrary.DBConnection.MyAdvantechGlobal
    Public strCPOOL As String = CenterLibrary.DBConnection.CurationPool
    Sub Main()
        Dim connMY As New SqlClient.SqlConnection(strMY)
        Dim connCPool As New SqlClient.SqlConnection(strCPOOL)

        Try
            Dim sql As String = " select a.COUNTRY, SUBSTRING(a.EMAIL_ADDRESS,0,200) as EMAIL, a.CREATED as DATE_TIME, 'Siebel' as SOURCE from SIEBEL_CONTACT a (nolock) where a.COUNTRY<>'' and a.EMAIL_ADDRESS like '%@%.%' and a.EMAIL_ADDRESS not like '%@advantech%.%' " + _
                                " union all " + _
                                " select t1.COUNTRY, t1.EMAIL, t1.DATE_TIME, t1.SOURCE from ( " + _
                                " select *, ROW_NUMBER() over (PARTITION BY t.EMAIL ORDER BY t.DATE_TIME) as row from ( " + _
                                " select ISNULL(RTRIM(LTRIM(a.COUNTRY_NAME)),'') as COUNTRY, SUBSTRING(RTRIM(LTRIM(a.CONTACT_EMAIL)),0,200) as EMAIL, a.SEND_DATE as DATE_TIME, 'ThankYou Letter' as SOURCE from CurationPool.dbo.B2BDIR_THANKYOU_LETTER_MASTER a (nolock) where a.COUNTRY_NAME <> '' " + _
                                " union " + _
                                " select a.COUNTRY, SUBSTRING(a.EMAIL,0,200), b.UPLOAD_DATE as DATE_TIME, 'List Upload' as SOURCE from CurationPool.dbo.LIST_DETAIL a (nolock) inner join CurationPool.dbo.LIST_MASTER b (nolock) on a.LIST_ID=b.ROW_ID where a.COUNTRY<>'' and a.EMAIL<>'' and a.EMAIL like '%@%.%' and a.EMAIL not like '%@advantech%.%' and a.UPLOAD_STATUS=1 " + _
                                " union " + _
                                " select a.COUNTRY, SUBSTRING(a.EMAIL,0,200), a.START as DATE_TIME, 'Live Chat' as SOURCE from CurationPool.dbo.RESOURCE_UPLOADED_DETAIL a (nolock) where a.COUNTRY<>'' and a.EMAIL<>'' and a.EMAIL like '%@%.%' and a.EMAIL not like '%@advantech%.%' " + _
                                " union  " + _
                                " select a.COUNTRY, SUBSTRING(a.EMAIL,0,200), a.TIMESTAMP as DATE_TIME, 'Curation Imported Log' as SOURCE from CurationPool.dbo.CURATION_ACTIVITY_IMPORTED_LOG a (nolock) where a.COUNTRY<>'' and a.EMAIL<>'' and a.EMAIL like '%@%.%' and a.EMAIL not like '%@advantech%.%' " + _
                                " ) as t ) as t1 where t1.row=1 "

            connMY.Open()

            Dim dt As New DataTable
            Dim apt As New SqlClient.SqlDataAdapter(sql, connMY)
            apt.SelectCommand.CommandTimeout = 600 * 1000
            apt.Fill(dt)

            Dim cmd As New SqlClient.SqlCommand("truncate table CurationPool.dbo.EMAIL_COUNTRY", connMY)
            cmd.CommandTimeout = 600 * 1000
            cmd.ExecuteNonQuery()

            connCPool.Open()
            Dim bk As New SqlClient.SqlBulkCopy(connCPool)
            bk.BulkCopyTimeout = 600 * 1000
            bk.DestinationTableName = "EMAIL_COUNTRY"
            bk.WriteToServer(dt)
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
            CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Sync EMAIL_COUNTRY Error", ex.ToString, True, "", "")
            connMY.Close() : connCPool.Close()
        End Try

        connMY.Close() : connCPool.Close()
    End Sub
    Public Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal Subject As String, ByVal Body As String, ByVal IsBodyHtml As Boolean)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml
        mySmtpClient = New Net.Mail.SmtpClient("172.20.0.76")
        mySmtpClient.Send(htmlMessage)
    End Sub
End Module
