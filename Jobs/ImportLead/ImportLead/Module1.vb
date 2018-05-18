Imports System.Data.SqlClient
Imports System.Data

Module Module1

    Sub Main()
        Dim strConnCPool As String = CenterLibrary.DBConnection.CurationPool
        Dim dt As New DataTable

        Dim apt As New SqlClient.SqlDataAdapter(
            " select t3.* from ( " +
            " select distinct t2.*, case when con.COUNTRY is not null and con.COUNTRY<>'' then con.COUNTRY else case when list.COUNTRY is not null and list.COUNTRY<>'' then list.COUNTRY else cura_act.COUNTRY end end as COUNTRY from (  " +
            " select distinct t1.EMAIL_ADDRESS,SUM(t1.eng_point) as point from (  " +
            " select distinct t.EMAIL_ADDRESS,SUM(t.eng_point) as eng_point from (   " +
            " select distinct EMAIL_ADDRESS, ACTIVITY_TYPE,  " +
            " case when URL like 'http%://www.advantech%.%/service/%' then 'http://www.advantech.com/service-dtos-sis-ags/'  " +
            " when URL like 'http%://www.advantech%.%/dtos/%' then 'http://www.advantech.com/service-dtos-sis-ags/'  " +
            " when URL like 'http%://www.advantech%.%/sis/%' then 'http://www.advantech.com/service-dtos-sis-ags/'  " +
            " when URL like 'http%://www.advantech%.%/ags/%' then 'http://www.advantech.com/service-dtos-sis-ags/' else URL end as URL,  " +
            " DATENAME(YEAR, CREATED_DATE)+'-'+DATENAME(MONTH, CREATED_DATE)+'-'+DATENAME(DAY, CREATED_DATE)+' '+DATENAME(HOUR, CREATED_DATE) as CREATED_DATE, eng_point from V_CURATION_ACTIVITY_ENGPOINT where IS_USED=0  " +
            " and EMAIL_ADDRESS not like '%@advantech%' and EMAIL_ADDRESS not like '%@advansus%' and EMAIL_ADDRESS<>'' and DESCRIPTION not in ('Idle','Pause')  " +
            " and CREATED_DATE between CONVERT(VARCHAR(10), dateadd(year,-1,GETDATE()), 111)+' 00:00:00.000' and CONVERT(VARCHAR(10), GETDATE(), 111)+' 23:59:59.999'  " +
            " and eng_point<100 and URL not like 'http%://member.advantech.com/yourcontactinformation.aspx?formid=%' and SOURCE_TYPE not in ('eDM','TracePart')) as t group by t.EMAIL_ADDRESS having SUM(t.eng_point)>0  " +
            " union all " +
            " select distinct t.EMAIL_ADDRESS,SUM(t.eng_point) as eng_point from (   " +
            " select distinct EMAIL_ADDRESS,  " +
            " case when URL like 'http%://www%.advantech%.%/service/%' then 'http://www.advantech.com/service-dtos-sis-ags/'  " +
            " when URL like 'http%://www%.advantech%.%/dtos/%' then 'http://www.advantech.com/service-dtos-sis-ags/'  " +
            " when URL like 'http%://www%.advantech%.%/sis/%' then 'http://www.advantech.com/service-dtos-sis-ags/'  " +
            " when URL like 'http%://www%.advantech%.%/ags/%' then 'http://www.advantech.com/service-dtos-sis-ags/' else URL end as URL,  " +
            " eng_point from V_CURATION_ACTIVITY_ENGPOINT a where a.IS_USED=0  " +
            " and a.EMAIL_ADDRESS not like '%@advantech%' and a.EMAIL_ADDRESS not like '%@advansus%' and a.EMAIL_ADDRESS<>'' and a.DESCRIPTION not in ('Idle','Pause')  " +
            " and a.CREATED_DATE between CONVERT(VARCHAR(10), dateadd(year,-1,GETDATE()), 111)+' 00:00:00.000' and CONVERT(VARCHAR(10), GETDATE(), 111)+' 23:59:59.999'  " +
            " and a.eng_point<100 and a.SOURCE_TYPE='eDM') as t  " +
            " group by t.EMAIL_ADDRESS  " +
            " union all  " +
            " select distinct t.EMAIL_ADDRESS, SUM(t.eng_point) as eng_point from (  " +
            " select distinct EMAIL_ADDRESS, replace(substring(URL,0,charindex('formid',url)+7+36),'https','http') as URL, eng_point  " +
            " from V_CURATION_ACTIVITY_ENGPOINT where IS_USED=0  " +
            " and EMAIL_ADDRESS not like '%@advantech%' and EMAIL_ADDRESS not like '%@advansus%' and EMAIL_ADDRESS<>''  " +
            " and CREATED_DATE between CONVERT(VARCHAR(10), dateadd(year,-1,GETDATE()), 111)+' 00:00:00.000' and CONVERT(VARCHAR(10), GETDATE(), 111)+' 23:59:59.999'  " +
            " and eng_point<100 and URL like 'http%://member.advantech.com/yourcontactinformation.aspx?formid=%' and SOURCE_TYPE not in ('eDM')) as t group by t.EMAIL_ADDRESS  " +
            " union all  " +
            " select distinct t.EMAIL_ADDRESS, SUM(t.eng_point) as eng_point from (  " +
            " select distinct EMAIL_ADDRESS, DATENAME(YEAR, CREATED_DATE)+'-'+DATENAME(MONTH, CREATED_DATE)+'-'+DATENAME(DAY, CREATED_DATE) as CREATED_DATE, eng_point  " +
            " from V_CURATION_ACTIVITY_ENGPOINT where IS_USED=0  " +
            " and EMAIL_ADDRESS not like '%@advantech%' and EMAIL_ADDRESS not like '%@advansus%' and EMAIL_ADDRESS<>''  " +
            " and CREATED_DATE between CONVERT(VARCHAR(10), dateadd(year,-1,GETDATE()), 111)+' 00:00:00.000' and CONVERT(VARCHAR(10), GETDATE(), 111)+' 23:59:59.999'  " +
            " and eng_point<100 and URL not like 'http%://member.advantech.com/yourcontactinformation.aspx?formid=%' and SOURCE_TYPE in ('TracePart')) as t group by t.EMAIL_ADDRESS  " +
            " ) as t1  " +
            " group by t1.EMAIL_ADDRESS having SUM(t1.eng_point) >= 100  " +
            " ) as t2  " +
            " inner join (select distinct z.EMAIL_ADDRESS from V_CURATION_ACTIVITY_ENGPOINT z where z.CREATED_DATE between DATEADD(day,-1,getdate()) and GETDATE() and z.eng_point>0 and z.IS_USED=0) t4 on t2.EMAIL_ADDRESS=t4.EMAIL_ADDRESS " +
            " left join MyAdvantechGlobal.dbo.SIEBEL_CONTACT con (nolock) on t2.EMAIL_ADDRESS=con.EMAIL_ADDRESS  " +
            " left join (select * from (select distinct z.EMAIL,z.COUNTRY,z1.UPLOAD_DATE,ROW_NUMBER() over (PARTITION BY z.EMAIL ORDER BY z1.UPLOAD_DATE desc) as row from LIST_DETAIL z (nolock) inner join LIST_MASTER z1 (nolock) on z.LIST_ID=z1.ROW_ID where z1.UPLOAD_STATUS=1) z2 where z2.row=1) list on t2.EMAIL_ADDRESS=list.EMAIL " +
            " left join (select * from (select distinct z.EMAIL,z.COUNTRY,z.TIMESTAMP,ROW_NUMBER() over (PARTITION BY z.EMAIL ORDER BY z.TIMESTAMP desc) as row from CURATION_ACTIVITY_IMPORTED_LOG z (nolock)) z2 where z2.row=1) cura_act on t2.EMAIL_ADDRESS=cura_act.EMAIL " +
            " where t2.EMAIL_ADDRESS not like '%@qq.com' " +
            " ) as t3  " +
            " where t3.COUNTRY is not null and t3.COUNTRY<>'' order by t3.point desc ",
            strConnCPool)
        apt.SelectCommand.CommandTimeout = 20 * 60
        Try
            apt.Fill(dt)
        Catch ex As Exception
            Console.Write(ex.ToString)
            'Console.Read()
        End Try


        Dim num_h As Integer = 0, num_w As Integer = 0
        Select Case dt.Rows.Count
            Case Is <= 3
                num_h = 3
            Case 4
                num_h = 2 : num_w = 3
            Case Else
                num_h = dt.Rows.Count / 5 : num_w = (dt.Rows.Count * 3 / 5) + num_h
        End Select

        Dim connCP As New SqlClient.SqlConnection(strConnCPool)

        For i As Integer = 0 To dt.Rows.Count - 1
            Try
                Dim ErrMsg As String = ""
                Dim ws As New IntApi_New.IntApi
                ws.UseDefaultCredentials = True : ws.Timeout = -1
                Dim quality As IntApi.LeadQuality
                If i + 1 <= num_h Then
                    quality = IntApi.LeadQuality.Hot
                ElseIf i + 1 > num_h AndAlso i + 1 <= num_w Then
                    quality = IntApi.LeadQuality.Warm
                Else
                    quality = IntApi.LeadQuality.Cool
                End If

                If ws.ImportLeadDaily(dt.Rows(i).Item("EMAIL_ADDRESS"), DateAdd(DateInterval.Year, -1, Now).ToString("yyyy-MM-dd"), Now.ToString("yyyy-MM-dd"), False, False, False, ErrMsg, quality, False, "") <> "" Then
                    Console.WriteLine(dt.Rows(i).Item("EMAIL_ADDRESS") + "  Success")
                Else
                    'Console.WriteLine(dt.Rows(i).Item("EMAIL_ADDRESS") + "  " + ErrMsg)
                    Dim cmd As New SqlClient.SqlCommand
                    If Not connCP.State = ConnectionState.Open Then connCP.Open()
                    cmd.Connection = connCP
                    cmd.CommandText = String.Format("insert into CURATION_SYSTEM_LOG (ActionName,InsertedTime,CustomerEmail,ErrorMsg) values ('ImportLeadDaily',getdate(),N'{0}',N'{1}')", dt.Rows(i).Item("EMAIL_ADDRESS").ToString, ErrMsg)
                    cmd.CommandTimeout = 5 * 60 * 1000
                    cmd.ExecuteNonQuery()
                End If
                Threading.Thread.Sleep(1 * 60 * 1000)
            Catch ex As Exception
                CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "ebiz.aeu@advantech.eu", "Import Daily Lead Failed(" + dt.Rows(i).Item("EMAIL_ADDRESS") + ")", ex.ToString, True, "", "")
            End Try
        Next
    End Sub

    Sub SendEmail( _
           ByVal SendTo As String, ByVal From As String, _
           ByVal Subject As String, ByVal Body As String, _
           ByVal IsBodyHtml As Boolean, _
           ByVal cc As String, _
           ByVal bcc As String, Optional ByVal NotifyOnFailure As Boolean = False)
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
        mySmtpClient = New Net.Mail.SmtpClient("172.20.1.12")
        Try
            mySmtpClient.Send(htmlMessage)
        Catch ex As System.Net.Mail.SmtpException
            System.Threading.Thread.Sleep(100)
            Try
                mySmtpClient.Send(htmlMessage)
            Catch ex1 As Exception
                Try
                    mySmtpClient = New Net.Mail.SmtpClient("172.20.1.12")
                    mySmtpClient.Send(htmlMessage)
                    htmlMessage = New Net.Mail.MailMessage("ebusiness.aeu@advantech.eu", "ebusiness.aeu@advantech.eu", "ACL SMTP Server Send Mail Failed", ex1.ToString)
                    mySmtpClient.Send(htmlMessage)
                Catch ex2 As Exception

                End Try
            End Try
        End Try
    End Sub
End Module
