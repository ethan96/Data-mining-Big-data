Imports System.Net

Module Module1
    Public strCPOOL As String = CenterLibrary.DBConnection.CurationPool
    Sub Main()
        Dim tables As String() = {"CURATION_ACTIVITY_TEMP_TRANSFER", "CURATION_ACTIVITY_TEMP_TRANSFER_2HR"}
        Try
            For Each table As String In tables
                Dim intApiWs As New CurationWS.IntApi
                Dim UnicaActDbConn As New SqlClient.SqlConnection(strCPOOL)
                Dim strModels As String = intApiWs.GetAllModelString()
                Dim strSelectUnicaActSql As String = _
                    " select a.ROW_ID, IsNull(a.DESCRIPTION,'') as DESCRIPTION, IsNull(a.COMMENT,'') as COMMENT, IsNull(a.URL,'') as URL, (select count(z.URL) from URL_MODEL_CACHE z where z.URL=a.URL) as nums " + _
                    " from " + table + " a " + _
                    " where a.PRODUCT_GROUP is null and a.INTERESTED_PRODUCT is null and a.MODEL_NO is null " + _
                    " and a.URL not in ('http://www.advantech.com','https://member.advantech.com','http://www.advantech.co.jp/','http://buy.advantech.com/Quotation/myquotation.aspx') " + _
                    " order by a.ROW_ID "
                Dim UnicaActDt As New DataTable
                Dim UnicaActApt As New SqlClient.SqlDataAdapter(strSelectUnicaActSql, UnicaActDbConn)
                UnicaActApt.Fill(UnicaActDt)
                For Each actRow As DataRow In UnicaActDt.Rows
                    'Update Model No by Description and Comment
                    Dim arrMatchedModels() As String = {}
                    Dim strCommentDesc As String = actRow.Item("DESCRIPTION") + actRow.Item("COMMENT")
                    If Not String.IsNullOrEmpty(strCommentDesc) Then
                        arrMatchedModels = intApiWs.GetMatchedModelsFromText(strCommentDesc, strModels)
                    End If
                    Dim cmd As New SqlClient.SqlCommand("update " + table + " set model_no=@MN where ROW_ID=@RID", UnicaActDbConn)
                    cmd.Parameters.AddWithValue("RID", actRow.Item("ROW_ID"))
                    If UnicaActDbConn.State <> ConnectionState.Open Then UnicaActDbConn.Open()
                    If arrMatchedModels IsNot Nothing AndAlso arrMatchedModels.Count > 0 Then
                        cmd.Parameters.AddWithValue("MN", arrMatchedModels(0))
                        cmd.ExecuteNonQuery()
                        Console.WriteLine("update " + arrMatchedModels(0) + " for " + actRow.Item("ROW_ID"))
                    Else
                        Dim strUrl As String = actRow.Item("URL")
                        If strUrl = "" Then
                            Dim cmd1 As New SqlClient.SqlCommand("update " + table + " set model_no=@MN where ROW_ID=@RID", UnicaActDbConn)
                            cmd1.Parameters.AddWithValue("RID", actRow.Item("ROW_ID"))
                            cmd1.Parameters.AddWithValue("MN", "N/A")
                            cmd1.ExecuteNonQuery()
                            Console.WriteLine("N/A update for " + actRow.Item("ROW_ID"))
                        Else
                            Dim arrUrlCache As New ArrayList
                            Dim intCount As Integer = CInt(actRow.Item("nums"))
                            If intCount = 0 Then
                                'Update URL_MODEL_CACHE
                                If Not arrUrlCache.Contains(strUrl) Then
                                    Dim blValidUrl As Boolean = False
                                    Try
                                        Dim UObj As New System.Uri(strUrl)
                                        If strUrl.StartsWith("http://", StringComparison.CurrentCultureIgnoreCase) = False And strUrl.StartsWith("https://", StringComparison.CurrentCultureIgnoreCase) = False Then
                                            blValidUrl = False
                                        Else
                                            If String.IsNullOrEmpty(strUrl) Then
                                                blValidUrl = False
                                            Else
                                                blValidUrl = True
                                            End If
                                        End If
                                    Catch ex As UriFormatException
                                        blValidUrl = False
                                    End Try
                                    Dim doc As HtmlAgilityPack.HtmlDocument = Nothing
                                    If blValidUrl Then
                                        doc = Nothing
                                        If GetTargetHTMLContent(strUrl, doc) Then
                                            Dim strContent As String = doc.DocumentNode.InnerText
                                            If UnicaActDbConn.State <> ConnectionState.Open Then UnicaActDbConn.Open()
                                            Try
                                                arrMatchedModels = intApiWs.GetMatchedModelsFromText(strContent, strModels)
                                                If arrMatchedModels.Length > 0 Then
                                                    Dim dtNow As DateTime = Now
                                                    Dim insDt As DataTable = GetInsDt()
                                                    For Each m As String In arrMatchedModels
                                                        Dim insRow As DataRow = insDt.NewRow()
                                                        insRow.Item("URL") = strUrl : insRow.Item("MODEL_NO") = m : insRow.Item("CREATED_DATE") = dtNow
                                                        insDt.Rows.Add(insRow)
                                                        Console.WriteLine(strUrl + ": " + m)
                                                    Next
                                                    Dim bk As New SqlClient.SqlBulkCopy(UnicaActDbConn)
                                                    If UnicaActDbConn.State <> ConnectionState.Open Then UnicaActDbConn.Open()
                                                    bk.DestinationTableName = "URL_MODEL_CACHE"
                                                    bk.WriteToServer(insDt)
                                                Else
                                                    Dim cmd1 As New SqlClient.SqlCommand("update " + table + " set model_no=@MN where URL=@URL and (DESCRIPTION='' or DESCRIPTION is null) and (COMMENT='' or COMMENT is null)", UnicaActDbConn)
                                                    cmd1.Parameters.AddWithValue("URL", strUrl)
                                                    cmd1.Parameters.AddWithValue("MN", "N/A")
                                                    cmd1.ExecuteNonQuery()
                                                    Console.WriteLine("N/A update for " + actRow.Item("ROW_ID"))
                                                End If
                                                arrUrlCache.Add(strUrl)
                                            Catch ex As Exception
                                                Dim cmd1 As New SqlClient.SqlCommand("update " + table + " set model_no=@MN where ROW_ID=@RID", UnicaActDbConn)
                                                cmd1.Parameters.AddWithValue("RID", actRow.Item("ROW_ID"))
                                                cmd1.Parameters.AddWithValue("MN", "N/A")
                                                cmd1.ExecuteNonQuery()
                                                Console.WriteLine("N/A update for " + actRow.Item("ROW_ID"))
                                            End Try
                                        End If
                                    Else
                                        Dim cmd1 As New SqlClient.SqlCommand("update " + table + " set model_no=@MN where ROW_ID=@RID", UnicaActDbConn)
                                        cmd1.Parameters.AddWithValue("RID", actRow.Item("ROW_ID"))
                                        cmd1.Parameters.AddWithValue("MN", "N/A")
                                        cmd1.ExecuteNonQuery()
                                        Console.WriteLine("N/A update for " + actRow.Item("ROW_ID"))
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next

                UnicaActDbConn.Close()
            Next
            
            Console.WriteLine("ok")
        Catch ex As Exception
            CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "ebiz.aeu@advantech.eu", "Update Model No for UNICA Activity error", ex.ToString, True, "", "")
        End Try
    End Sub

    Public Function GetInsDt() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("URL") : dt.Columns.Add("MODEL_NO") : dt.Columns.Add("CREATED_DATE", GetType(DateTime))
        Return dt
    End Function

    Function GetTargetHTMLContent(ByVal URL As String, ByRef doc As HtmlAgilityPack.HtmlDocument) As Boolean
        Try
            Dim client As New WebDownload
            doc = New HtmlAgilityPack.HtmlDocument
            Dim ms As IO.MemoryStream = Nothing
            Try
                ms = New IO.MemoryStream(client.DownloadData(URL))
            Catch ex As Exception
                Return False
            End Try
            'Console.WriteLine(ms.Length.ToString())
            If ms.Length < 200000 Then
                doc.Load(ms, Text.Encoding.GetEncoding("utf-8")) : Return True
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal Subject As String, ByVal Body As String, ByVal IsBodyHtml As Boolean)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml
        mySmtpClient = New Net.Mail.SmtpClient("172.20.0.76")
        mySmtpClient.Send(htmlMessage)
    End Sub

    Public Class WebDownload : Inherits WebClient
        Private _timeout As Integer

        Public Property Timeout() As Integer
            Get
                Return _timeout
            End Get
            Set(value As Integer)
                _timeout = value
            End Set
        End Property

        Public Sub New()
            Me._timeout = 60000
        End Sub

        Public Sub New(timeout As Integer)
            Me._timeout = timeout
        End Sub

        Protected Overrides Function GetWebRequest(address As Uri) As WebRequest
            Dim result = MyBase.GetWebRequest(address)
            result.Timeout = Me._timeout
            Return result
        End Function
    End Class
End Module
