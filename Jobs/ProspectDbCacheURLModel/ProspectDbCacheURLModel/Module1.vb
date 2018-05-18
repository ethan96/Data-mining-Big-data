Module Module1
    Public strMyConn As String = CenterLibrary.DBConnection.MyAdvantechGlobal
    Public strProsDbConn As String = CenterLibrary.DBConnection.CurationPool
    Sub Main()
        Go()
        Console.WriteLine("Done")
        Console.Read()
    End Sub

    Public Sub Go()
        Dim strModels As String = Utility.GetAllModelStrings()
        Dim PosDbConn As New SqlClient.SqlConnection(strProsDbConn)
        PosDbConn.Open()
        Dim strSql As String = _
            " select distinct a.URL, (select count(z.URL) from URL_MODEL_CACHE z where z.URL=a.URL) as nums from CURATION_ACTIVITY a where a.MODEL_NO is null and a.URL like 'http%//%.advantech.%' and a.CREATED_DATE between dateadd(month,-3,getdate()) and getdate() and a.URL not in " + _
            " ('http://www.advantech.com','https://member.advantech.com','http://www.advantech.co.jp/','http://buy.advantech.com/Quotation/myquotation.aspx') order by a.URL "
        Dim dt As New DataTable
        Dim apt As New SqlClient.SqlDataAdapter(strSql, PosDbConn)
        apt.Fill(dt)
        For Each r As DataRow In dt.Rows
            Dim strUrl As String = r.Item("URL")
            Dim intCount As Integer = CInt(r.Item("nums"))
            If intCount = 0 Then
                Dim doc As HtmlAgilityPack.HtmlDocument = Nothing
                If GetTargetHTMLContent(strUrl, doc) Then
                    Dim strContent As String = doc.DocumentNode.InnerText
                    Dim mAry As ArrayList = Utility.GetMatchedModelsFromText(strContent, strModels)
                    If mAry.Count > 0 Then
                        Dim dtNow As DateTime = Now
                        Dim insDt As DataTable = GetInsDt()
                        For Each m As String In mAry
                            Dim insRow As DataRow = insDt.NewRow()
                            insRow.Item("URL") = strUrl : insRow.Item("MODEL_NO") = m : insRow.Item("CREATED_DATE") = dtNow
                            insDt.Rows.Add(insRow)
                            Console.WriteLine(strUrl + ": " + m)
                        Next
                        Dim bk As New SqlClient.SqlBulkCopy(PosDbConn)
                        If PosDbConn.State <> ConnectionState.Open Then PosDbConn.Open()
                        bk.DestinationTableName = "URL_MODEL_CACHE"
                        bk.WriteToServer(insDt)
                    End If
                End If
            End If
           
        Next
        PosDbConn.Close()
    End Sub

    Public Function GetInsDt() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("URL") : dt.Columns.Add("MODEL_NO") : dt.Columns.Add("CREATED_DATE", GetType(DateTime))
        Return dt
    End Function

    Function GetTargetHTMLContent(ByVal URL As String, ByRef doc As HtmlAgilityPack.HtmlDocument) As Boolean
        Dim client As New Net.WebClient
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
    End Function

End Module
