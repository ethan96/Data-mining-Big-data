Imports System.Data.SqlClient

Module Module1
    Public strMyGlobal As String = CenterLibrary.DBConnection.CurationPool
    'Public strAEUSQL As String = "Data Source=aeusql-hq;Initial Catalog=MyAdvantech;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    'Frank 20140825 eStore DB has a hardware issue, so target db connnection string switchs to 172.21.1.20 temporarily
    'Public strSSO As String = "Data Source=172.21.1.18;Initial Catalog=membership;Persist Security Info=True;User ID=myApp;Password=my@pp;async=true;Connect Timeout=300;pooling='true'"
    Public strSSO As String = CenterLibrary.DBConnection.Membership

    Sub Main()
        Try
            Dim dtSSO As New DataTable
            Using conn As New SqlClient.SqlConnection(strSSO)
                conn.Open()
                Dim apt As New SqlClient.SqlDataAdapter(" select distinct email_addr, case User_Status when -1 then 0 else 1 end as user_status, country " +
                                " from membership.dbo.member " +
                                " where email_addr like '%@%.%' ", conn)
                apt.Fill(dtSSO)
            End Using

            Using conn As New SqlClient.SqlConnection(strMyGlobal)
                conn.Open()
                Dim bk As New SqlClient.SqlBulkCopy(conn)
                bk.DestinationTableName = "SSO_MEMBER"
                bk.WriteToServer(dtSSO)

                Dim cmd As New SqlClient.SqlCommand("delete from SSO_MEMBER where FLAG=1", conn)
                cmd.CommandTimeout = 600 * 1000
                cmd.ExecuteNonQuery()

                cmd = New SqlCommand("update SSO_MEMBER set FLAG=1", conn)
                cmd.CommandTimeout = 600 * 1000
                cmd.ExecuteNonQuery()
            End Using



            'Dim sql As String = "with source_data as ( " +
            '                    " select distinct email_addr, case User_Status when -1 then 0 else 1 end as user_status, country " +
            '                    " from [172.21.1.20].membership.dbo.member " +
            '                    " where email_addr like '%@%.%'  " +
            '                    " ) " +
            '                    " MERGE INTO SSO_MEMBER A Using source_data B  " +
            '                    " On A.email_addr=B.email_addr  " +
            '                    " WHEN NOT MATCHED THEN  " +
            '                    " INSERT (email_addr,user_status,country) VALUES " +
            '                    " (B.email_addr,B.user_status,B.country) " +
            '                    " WHEN MATCHED AND A.user_status <> B.user_status THEN " +
            '                    " UPDATE SET A.user_status=B.user_status; "
            ''Dim conn As New SqlClient.SqlConnection(strMyGlobal)
            'conn.Open()
            'Dim cmd As New SqlClient.SqlCommand(sql, conn)
            'cmd.CommandTimeout = 6000 * 1000
            'cmd.ExecuteNonQuery()
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
            Try
                CenterLibrary.MailUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SSO member failed", ex.ToString(), False, "", "")
            Catch ex2 As Exception
                Console.WriteLine(ex2.ToString())
            End Try
        End Try
        'Dim intDiff As Integer = 7
        'Dim strSql As String = _
        '    " select distinct email_addr, case User_Status when -1 then 0 else 1 end as user_status, country  " + _
        '    " from member  " + _
        '    " where email_addr like '%@%.%'  " + _
        '    " and (DATE_REGISTERED between GETDATE()-" + intDiff.ToString() + " and GETDATE()+1 or DATE_LAST_CHANGED between GETDATE()-" + intDiff.ToString() + " and GETDATE()+1) " + _
        '    " order by email_addr "
        'Try
        '    Dim dt As DataTable = dbGetDataTable(strSSO, strSql)
        '    If dt.Rows.Count > 0 Then
        '        Console.WriteLine("total " + dt.Rows.Count.ToString() + " member(s) to update")
        '        Dim sqlMyLocal As New SqlConnection(strMyGlobal), dbCmd As SqlCommand = Nothing
        '        sqlMyLocal.Open()
        '        For Each r As DataRow In dt.Rows
        '            dbCmd = New SqlCommand("delete from SSO_MEMBER where email_addr='" + Replace(r.Item("email_addr").ToString(), "'", "''") + "'", sqlMyLocal)
        '            dbCmd.ExecuteNonQuery()
        '            Console.WriteLine("deleted " + r.Item("email_addr"))
        '        Next
        '        sqlMyLocal.Close()
        '        Dim bk As New SqlBulkCopy(strMyGlobal)
        '        bk.DestinationTableName = "SSO_MEMBER"
        '        bk.WriteToServer(dt)

        '        Console.WriteLine("updated in MyGlobal CurationPool")

        '        'sqlMyLocal = New SqlConnection(strAEUSQL)
        '        'sqlMyLocal.Open()
        '        'For Each r As DataRow In dt.Rows
        '        '    dbCmd = New SqlCommand("delete from SSO_MEMBER where email_addr='" + Replace(r.Item("email_addr").ToString(), "'", "''") + "'", sqlMyLocal)
        '        '    dbCmd.ExecuteNonQuery()
        '        '    Console.WriteLine("deleted " + r.Item("email_addr"))
        '        'Next
        '        'sqlMyLocal.Close()

        '        'bk = New SqlBulkCopy(strAEUSQL)
        '        'bk.DestinationTableName = "SSO_MEMBER"
        '        'bk.WriteToServer(dt)

        '        'Console.WriteLine("updated in aeusql")

        '    End If
        'Catch ex As Exception
        '    Console.WriteLine(ex.ToString())
        '    Try
        '        SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SSO member failed", ex.ToString(), False)
        '    Catch ex2 As Exception
        '        Console.WriteLine(ex2.ToString())
        '    End Try
        '    Threading.Thread.Sleep(10000)
        'End Try

        Console.WriteLine("done")
        Threading.Thread.Sleep(5000)
    End Sub

    Public Function dbGetDataTable(ByVal ConnectionName As String, ByVal strSqlCmd As String) As DataTable
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

    Public Sub SendEmail( _
         ByVal SendTo As String, ByVal From As String, _
         ByVal Subject As String, ByVal Body As String, _
         ByVal IsBodyHtml As Boolean)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml

        mySmtpClient = New Net.Mail.SmtpClient("172.20.0.76")
        Try
            mySmtpClient.Send(htmlMessage)
        Catch ex As System.Net.Mail.SmtpException
            System.Threading.Thread.Sleep(100)
            mySmtpClient.Send(htmlMessage)
        End Try
    End Sub

End Module
