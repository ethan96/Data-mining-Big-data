Imports NCrawler.Interfaces

Public Class CrawlSite
    Public gCounter As GlobalCounter
    Public gdt As DataTable = Nothing ', ErrLnkDt As DataTable = Nothing
    Private _strMyGlob As String = "Data Source=aclsql6\sql2008r2;Initial Catalog=MyAdvantechGlobal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    Public _TargetUrl As String ', _RowId As String
    Public _ErrMsg As String
    Public _FinFlag As Boolean, _MaxThreads As Integer, _MaxCDepth As Integer, _AppName As String

    Public ReadOnly Property ErrorMessage As String
        Get
            If _ErrMsg <> String.Empty Then Return _ErrMsg
            Return ""
        End Get
    End Property

    Public Sub New(ByVal SiteUrl As String, ByVal MaxThreads As Integer, ByVal MaxDepth As Integer, ByRef gc As GlobalCounter, ByVal AppName As String)
        _TargetUrl = SiteUrl : _ErrMsg = "" : _FinFlag = False : _MaxThreads = MaxThreads : _MaxCDepth = MaxDepth
        gCounter = gc : _AppName = AppName
    End Sub

    Sub StartCrawl()
        SyncLock GetType(GlobalCounter)
            gCounter.c = gCounter.c + 1
        End SyncLock
        gdt = New DataTable("CrawlDt") 
        With gdt.Columns
            .Add("KEYID") : .Add("Url") : .Add("ContentType")
            .Add("LastModified", GetType(Date)) : .Add("ResponseUri")
            '.Add("Server")
            .Add("Title") : .Add("Text") : .Add("Meta_Description") : .Add("Meta_Keywords") : .Add("Depth", GetType(Integer))
            .Add("Crawl_Time", GetType(DateTime)) : .Add("APPNAME") : .Add("GOOGLE_PAGERANK", GetType(Integer))
        End With
        'With ErrLnkDt.Columns
        '    .Add("Uri") : .Add("Referrer") : .Add("Detect_Time", GetType(DateTime))
        'End With
        Try
            Dim tu As New System.Uri(_TargetUrl)
        Catch ex As Exception
            _ErrMsg = ex.ToString()
            _FinFlag = True : Exit Sub
        End Try
        Dim c As New NCrawler.Crawler(New System.Uri(_TargetUrl), New NCrawler.HtmlProcessor.HtmlDocumentProcessor(), New DumperStep(gdt, _TargetUrl))
        'c.ConnectionTimeout = New TimeSpan(0, 0, 60) : c.ConnectionReadTimeout = New TimeSpan(0, 0, 60)
        c.MaximumThreadCount = _MaxThreads
        c.MaximumCrawlDepth = _MaxCDepth
        c.ExcludeFilter = New NCrawler.Services.RegexFilter() { _
            New NCrawler.Services.RegexFilter( _
                New Text.RegularExpressions.Regex("(\.jpg|\.css|\.js|\.gif|\.jpeg|\.png|\.ico|\.swf|\.flv|\.wmv|\.mp3|\.mpg|\.mpeg|\.rm|\.rmvb|\.mkv|\.mka|\.avi)", _
                Text.RegularExpressions.RegexOptions.Compiled Or _
                Text.RegularExpressions.RegexOptions.CultureInvariant Or _
                Text.RegularExpressions.RegexOptions.IgnoreCase))}
        AddHandler c.CrawlFinished, AddressOf CrawlFinished
        AddHandler c.PipelineException, AddressOf PipelineException
        c.Crawl()
    End Sub


    Private Sub CrawlFinished(ByVal sender As Object, ByVal e As NCrawler.Events.CrawlFinishedEventArgs)
        Try
            If gdt.Rows.Count > 0 Then
                Dim conn As New SqlClient.SqlConnection(_strMyGlob)
                Dim dbCmd As New SqlClient.SqlCommand("delete from MY_WEB_SEARCH where url='" + Me._TargetUrl + "'", conn)
                dbCmd.CommandTimeout = 20 * 60
                conn.Open()
                dbCmd.ExecuteNonQuery()
                conn.Close()
                For Each r As DataRow In gdt.Rows
                    r.Item("KEYID") = System.Guid.NewGuid.ToString()
                    If r.Item("ContentType").ToString.Length > 500 Then r.Item("ContentType") = Left(r.Item("ContentType"), 500)
                    'If r.Item("OriginalReferrerUrl").ToString.Length > 1500 Then r.Item("OriginalReferrerUrl") = Left(r.Item("OriginalReferrerUrl"), 1500)
                    'If r.Item("OriginalUrl").ToString.Length > 1500 Then r.Item("OriginalUrl") = Left(r.Item("OriginalUrl"), 1500)
                    'If r.Item("Referrer").ToString.Length > 1500 Then r.Item("Referrer") = Left(r.Item("Referrer"), 1500)
                    If r.Item("ResponseUri").ToString.Length > 1500 Then r.Item("ResponseUri") = Left(r.Item("ResponseUri"), 1500)
                    'If r.Item("Server").ToString.Length > 500 Then r.Item("Server") = Left(r.Item("Server"), 500)
                    If r.Item("Title").ToString.Length > 500 Then r.Item("Title") = Left(r.Item("Title"), 500)
                    If r.Item("Meta_Description").ToString.Length > 500 Then r.Item("Meta_Description") = Left(r.Item("Meta_Description"), 500)
                    If r.Item("Meta_Keywords").ToString.Length > 500 Then r.Item("Meta_Keywords") = Left(r.Item("Meta_Keywords"), 500)
                    r.Item("APPNAME") = Me._AppName
                    r.Item("GOOGLE_PAGERANK") = -1
                    If Date.TryParse(r.Item("LastModified"), Now) = False Then r.Item("LastModified") = Date.MinValue
                    Threading.Thread.Sleep(10)
                Next
                Dim bk As New SqlClient.SqlBulkCopy(_strMyGlob)
                bk.DestinationTableName = "MY_WEB_SEARCH"
                bk.WriteToServer(gdt)
                
            End If
        Catch ex As Exception
            _ErrMsg += "CrawlFinished exception:" + ex.ToString()
        End Try
        If _ErrMsg <> "" Then
            Try
                Dim htmlMessage As New Net.Mail.MailMessage("tc.chen@advantech.com.tw", "tc.chen@advantech.com.tw", "Crawl " + Me._TargetUrl + " error", _ErrMsg)
                Dim mySmtpClient As New Net.Mail.SmtpClient("172.21.34.21")
                mySmtpClient.Send(htmlMessage)
            Catch ex As Exception
            End Try
        Else
        End If
        _FinFlag = True
        SyncLock GetType(GlobalCounter)
            gCounter.c = gCounter.c - 1
        End SyncLock
    End Sub

    Public Class DumperStep
        Implements IPipelineStep
        Private _ProcDt As DataTable, _Url As String ', _errDt As DataTable
        Public Sub New(ByRef dt As DataTable, ByVal CrawlingUrl As String)
            _ProcDt = dt : _Url = CrawlingUrl
        End Sub
        Private Shared Function ToUnicodeString(ByVal str As String) As String
            If str Is Nothing Then Return ""
            Return str
            'Return str
            'Dim byteArray() As Byte = Text.Encoding.Unicode.GetBytes(str)
            'Return Text.Encoding.Unicode.GetString(byteArray)
        End Function
        Public Sub Process(ByVal crawler As NCrawler.Crawler, ByVal propertyBag As NCrawler.PropertyBag) Implements NCrawler.Interfaces.IPipelineStep.Process
            'Threading.Thread.Sleep(100)
            Dim r As DataRow = Nothing, tmpMetaDesc As String = "", tmpMetaKeywords As String = ""
            SyncLock GetType(DataTable)
                r = _ProcDt.NewRow()
            End SyncLock
            If propertyBag.StatusCode = Net.HttpStatusCode.OK Then
                If IsPdfContent(propertyBag.ContentType) Then
                    Using input As IO.Stream = propertyBag.GetResponse.Invoke
                        Dim pdfReader As iTextSharp.text.pdf.PdfReader = Nothing
                        Try
                            pdfReader = New iTextSharp.text.pdf.PdfReader(input)
                            Dim title As Object = pdfReader.Info.Item("Title")
                            If Not title IsNot Nothing Then
                                Dim pdfTitle As String = Convert.ToString(title, Globalization.CultureInfo.InvariantCulture).Trim
                                If Not pdfTitle = String.Empty Then
                                    propertyBag.Title = pdfTitle
                                End If
                            End If
                            Dim sb As New Text.StringBuilder
                            Dim p As Integer = 1
                            Do While (p <= pdfReader.NumberOfPages)
                                Dim pageBytes As Byte() = pdfReader.GetPageContent(p)
                                If Not pageBytes Is Nothing AndAlso pageBytes.Length > 0 Then
                                    Dim token As New iTextSharp.text.pdf.PRTokeniser(pageBytes)
                                    Try
                                        Do While token.NextToken
                                            Dim tknType As Integer = token.TokenType
                                            Dim tknValue As String = token.StringValue
                                            If (tknType = 2) Then
                                                sb.Append(token.StringValue)
                                                sb.Append(" ")
                                            ElseIf ((tknType = 1) AndAlso (tknValue = "-600")) Then
                                                sb.Append(" ")
                                            ElseIf ((tknType = 10) AndAlso (tknValue = "TJ")) Then
                                                sb.Append(" ")
                                            End If
                                        Loop
                                    Catch ex As Exception

                                    End Try

                                End If
                                p += 1
                            Loop
                            propertyBag.Text = sb.ToString
                        Finally
                            If pdfReader IsNot Nothing Then pdfReader.Close()
                        End Try
                    End Using
                End If
                'r.Item("Html") = ""
                If IsHtml(propertyBag.ContentType) Then
                    Try
                        Dim doc As New HtmlAgilityPack.HtmlDocument()
                        Dim o As IO.Stream = propertyBag.GetResponse.Invoke
                        Dim loadFlag As Boolean = False
                        Try
                            doc.Load(o, True)
                            loadFlag = True
                        Catch ex As System.OutOfMemoryException
                            System.GC.Collect()
                        End Try
                        If loadFlag Then
                            'r.Item("Html") = doc.DocumentNode.OuterHtml
                            'Console.WriteLine(doc.DocumentNode.OuterHtml)
                            Dim mns As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//meta")
                            If mns IsNot Nothing AndAlso mns.Count > 0 Then
                                For Each mn As HtmlAgilityPack.HtmlNode In mns
                                    'Console.WriteLine(String.Format("name:{0}|content:{1}", mn.GetAttributeValue("name", ""), mn.GetAttributeValue("content", "")))
                                    If mn.GetAttributeValue("name", "").ToUpper() = "DESCRIPTION" Then
                                        tmpMetaDesc = mn.GetAttributeValue("content", "")
                                    Else
                                        If mn.GetAttributeValue("name", "").ToUpper() = "KEYWORDS" Then
                                            tmpMetaKeywords = mn.GetAttributeValue("content", "")
                                        End If
                                    End If
                                Next
                                'Console.Read()
                            Else
                                'Console.WriteLine("No Meta")
                            End If
                        End If

                    Catch ex As Exception
                        Console.WriteLine("Get meta error:" + ex.ToString())
                        'Console.Read()
                    End Try
                End If

                With propertyBag
                    r.Item("KEYID") = "" : r.Item("Url") = ToUnicodeString(_Url)
                    r.Item("ContentType") = ToUnicodeString(.ContentType)
                    r.Item("LastModified") = .LastModified
                    'If .OriginalReferrerUrl IsNot Nothing Then
                    '    r.Item("OriginalReferrerUrl") = ToUnicodeString(.OriginalReferrerUrl.ToString())
                    'Else
                    '    r.Item("OriginalReferrerUrl") = ""
                    'End If
                    'r.Item("OriginalUrl") = ToUnicodeString(.OriginalUrl) : r.Item("Referrer") = ToUnicodeString(.Referrer.Uri.ToString())
                    r.Item("ResponseUri") = ToUnicodeString(.ResponseUri.ToString()) 'r.Item("Server") = ToUnicodeString(.Server)
                    r.Item("Title") = ToUnicodeString(.Title) : r.Item("Text") = ToUnicodeString(.Text)
                    r.Item("Meta_Description") = tmpMetaDesc : r.Item("Meta_Keywords") = tmpMetaKeywords
                    r.Item("Depth") = .Step.Depth
                    r.Item("Crawl_Time") = Now
                End With
                If r.Item("Text").ToString() <> "" Then
                    SyncLock GetType(DataTable)
                        _ProcDt.Rows.Add(r)
                    End SyncLock
                End If
            Else
                'Try
                '    Dim er As DataRow = Nothing
                '    SyncLock GetType(DataTable)
                '        er = _errDt.NewRow()
                '    End SyncLock
                '    With er
                '        er.Item("Uri") = propertyBag.ResponseUri.ToString() : er.Item("Referrer") = propertyBag.OriginalReferrerUrl.ToString()
                '        er.Item("Detect_Time") = Now
                '        'dt.Rows.Add(r)
                '    End With
                '    Console.WriteLine("err link detected:" + er.Item("Uri"))
                '    SyncLock GetType(DataTable)
                '        _errDt.Rows.Add(er)
                '    End SyncLock
                'Catch ex As Exception
                '    Console.WriteLine("Error while getting error link:" + ex.ToString())
                'End Try
               
            End If
        End Sub

        Private Shared Function IsPdfContent(ByVal contentType As String) As Boolean
            Return contentType.StartsWith("application/pdf", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsHtml(ByVal contentType As String) As Boolean
            Return contentType.StartsWith("text/html", StringComparison.OrdinalIgnoreCase)
        End Function

    End Class

    Private Sub PipelineException(ByVal sender As Object, ByVal e As NCrawler.Events.PipelineExceptionEventArgs)
        _ErrMsg += "PipelineException:" + e.Exception.ToString() + vbCrLf
    End Sub

End Class
