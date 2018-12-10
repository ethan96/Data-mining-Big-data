Module Module1
    Class AdvSite
        Public siteUrl As String, siteName As String, TillDepth As Integer
        Sub New(ByVal u As String, ByVal n As String, ByVal d As Integer)
            siteUrl = u : siteName = n : TillDepth = d
        End Sub
    End Class

    Dim _strMyGlob As String = "Data Source=aclsql6\sql2008r2;Initial Catalog=MyAdvantechGlobal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    Sub Main()
        Try
            Dim sites() As AdvSite = { _
                New AdvSite("http://webaccess.advantech.com/", "Advantech US", 7), _
                New AdvSite("http://www.advantech-innocore.com", "Advantech Innocore", 7), _
                 New AdvSite("http://www.advantech.com", "Advantech US", 7), _
               New AdvSite("http://buy.advantech.com", "eStore US", 7), _
               New AdvSite("http://support.advantech.com", "Support", 4)}
            Dim GCo As New GlobalCounter
            For Each site As AdvSite In sites
                Dim targetUrl As String = site.siteUrl
                If Not targetUrl.StartsWith("http") Then targetUrl = "http://" + targetUrl
                If IsValidUrlFormat(targetUrl) Then
                    Console.WriteLine("start crawling " + targetUrl)
                    Dim cw As New CrawlSite(targetUrl, 10, site.TillDepth, GCo, site.siteName)
                    cw.StartCrawl()
                    If cw.ErrorMessage <> "" Then
                        Console.WriteLine(targetUrl + " errmsg:" + cw.ErrorMessage)
                    End If
                    Console.WriteLine("del:" + cw._TargetUrl)
                    'Console.Read()
                End If
                'Exit For
            Next
        Catch ex As Exception
            Try
                Dim htmlMessage As New Net.Mail.MailMessage("tc.chen@advantech.com.tw", "tc.chen@advantech.com.tw", "Global MA Crawl ADV Sites Error", ex.ToString())
                Dim mySmtpClient As New Net.Mail.SmtpClient("172.21.34.21")
                mySmtpClient.Send(htmlMessage)
            Catch ex2 As Exception

            End Try
        End Try
        dbGetDataTable("delete from MY_WEB_SEARCH where Title='Advantech - Page Not Found'")
        dbGetDataTable("delete from MY_WEB_SEARCH where ResponseUri like 'http%//member.advantech.com/login.aspx%'")
        Console.WriteLine("Completed")
        'System.Diagnostics.Process.Start("E:\Scheduled_Programs\GetAdvWebPageRank\bin\Debug\GetAdvWebPageRank.exe")
        'Console.Read()
    End Sub

    Public Function dbGetDataTable(ByVal strSqlCmd As String) As DataTable
        Dim g_adoConn As New SqlClient.SqlConnection(_strMyGlob)
        Dim dt As New DataTable
        Dim da As New SqlClient.SqlDataAdapter(strSqlCmd, g_adoConn)
        da.SelectCommand.CommandTimeout = 20 * 60
        Try
            da.Fill(dt)
        Catch ex As Exception
            g_adoConn.Close() : Throw New Exception(ex.ToString() + vbTab + "sql:" + strSqlCmd)
        End Try
        g_adoConn.Close() : g_adoConn = Nothing
        Return dt
    End Function

    Private Function IsValidUrlFormat(ByVal email As String) As Boolean
        Dim reg As String = "(http:\/\/([\w.]+\/?)\S*)", reg2 As String = "(https:\/\/([\w.]+\/?)\S*)"
        Dim options As Text.RegularExpressions.RegexOptions = Text.RegularExpressions.RegexOptions.Singleline
        If Text.RegularExpressions.Regex.Matches(email, reg, options).Count = 0 Then
            If Text.RegularExpressions.Regex.Matches(email, reg2, options).Count = 0 Then
                Return False
            Else
                Return True
            End If

        Else
            Return True
        End If
    End Function

End Module
