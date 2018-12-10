Imports System.Security.Cryptography.X509Certificates
Imports Google.Apis.Auth.OAuth2
Imports Google.Apis.Plus.v1
Imports Google.Apis.Plus.v1.Data
Imports Google.Apis.Services
Imports Google.Apis.Analytics.v3
Imports Google.Apis.Analytics.v3.Data
Imports System.Data.SqlClient

Module Module1
    Public CurationPoolConnStr As String = "Data Source=ACLSQL6\SQL2008R2;Initial Catalog=CurationPool;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;Application Name=MyAdvantech;Failover Partner=ACLSQL7\SQL2008R2;async=true;Connect Timeout=300;pooling='true'"
    Public MyLocal As String = "Data Source=aclecampaign2\MATEST;Initial Catalog=MyLocal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;Application Name=MyAdvantech;Connect Timeout=300;pooling='true'"
    Public SearchMetrics As String = "ga:entrances,ga:pageviews"
    Public SearchDimension As String = "ga:searchKeyword,ga:country,ga:medium,ga:dateHour,ga:city,ga:landingPagePath,ga:networkDomain"
    Public TrafficMetrics As String = "ga:pageviews,ga:bounces"
    Public TrafficDimension As String = "ga:country,ga:language,ga:dateHour,ga:city,ga:landingPagePath,ga:networkDomain,ga:campaign"
    Public Sub Main(args As String())
        GATrafficCampaign()
        GASearchSourceMedium()
        Console.WriteLine("done")
        'Console.Read()
        Exit Sub
        '
        'URL for accessing Google Developers Console
        'https://console.developers.google.com/project/127038164893/apiui/credential?authuser=1
        'Google Analytics Core Reporting API with C# 
        'http://www.daimto.com/googleAnalytics-core-csharp/
        'Dim _GoogleAccount As String = "Advantech.Analytics@gmail.com"
        'Dim _GooglePassword As String = "MontBlanc201508"

        Dim dirInfo As New IO.DirectoryInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)

        ' view and manage your Google Analytics data
        Dim scopes As String() = New String() {AnalyticsService.Scope.AnalyticsReadonly}
        ' View Google Analytics data      
        Dim keyFilePath As String = dirInfo.Parent.Parent.Parent.FullName + "\key.p12"
        ' found in developer console
        Dim serviceAccountEmail As String = "127038164893-9qlqfak1plsd7uv9t5jmr9l12dkj86ub@developer.gserviceaccount.com"
        ' found in developer console
        'loading the Key file
        Dim certificate = New X509Certificate2(keyFilePath, "notasecret", X509KeyStorageFlags.Exportable)

        Dim Initializer1 As New ServiceAccountCredential.Initializer(serviceAccountEmail)
        Initializer1.Scopes = scopes : Initializer1.FromCertificate(certificate)

        Dim credential As New ServiceAccountCredential(Initializer1)

        Dim Initializer2 As New BaseClientService.Initializer()
        Initializer2.HttpClientInitializer = credential : Initializer2.ApplicationName = "API Projects"

        Dim service As New AnalyticsService(Initializer2)

        Dim GetDataTypes() As String = {"Traffic", "Search"}

        For Each gType In GetDataTypes
            Dim Dimension As String = "", Metrics As String = "", writeToTable As String = "", intDayDiff As Integer = 0
            Select Case gType
                Case "Traffic"
                    Dimension = TrafficDimension : Metrics = TrafficMetrics : writeToTable = "GOOGLE_GA_TRAFFIC" : intDayDiff = 1
                Case "Search"
                    Dimension = SearchDimension : Metrics = SearchMetrics : writeToTable = "GOOGLE_GA_SEARCH_RAW" : intDayDiff = 6
            End Select
            Dim FromDate As Date = DateAdd(DateInterval.Day, -4, Now)
            Dim iteration As Integer = DateDiff(DateInterval.Day, FromDate, DateAdd(DateInterval.Day, -1, Now)) / intDayDiff
            For idxDays As Integer = 0 To iteration
                Dim StartDate As New Date(FromDate.Year, FromDate.Month, FromDate.Day), EndDate As Date = DateAdd(DateInterval.Day, intDayDiff, StartDate)
                If DateDiff(DateInterval.Day, Now, StartDate) >= 0 Then Exit For
                Console.WriteLine("{0}~{1} of {2}", StartDate.ToString("yyyy-MM-dd"), EndDate.ToString("yyyy-MM-dd"), gType)
                Dim _datatable As New DataTable
                Dim request As DataResource.GaResource.GetRequest = service.Data.Ga.Get("ga:9115484", StartDate.ToString("yyyy-MM-dd"), EndDate.ToString("yyyy-MM-dd"), Metrics)
                request.Dimensions = Dimension : request.MaxResults = 10000 : request.StartIndex = 1
                Dim result As GaData = request.Execute()
                'Dim allRows As New List(Of String)
                '''/ Loop through until we arrive at an empty page
                While result.Rows IsNot Nothing
                    'Add the rows to the final list
                    'allRows.AddRange(result.Rows)
                    ' We will know we are on the last page when the next page token is
                    ' null.
                    ' If this is the case, break.
                    If result.NextLink Is Nothing Then
                        If request.MaxResults <= 1 Then
                            'Console.WriteLine("no more rows")
                            Exit While
                        Else
                            request.MaxResults /= 10
                        End If

                    Else
                        Dim dtResult As New DataTable
                        For Each headers In result.ColumnHeaders
                            dtResult.Columns.Add(headers.Name)
                        Next

                        For Each row In result.Rows
                            Dim nr As DataRow = dtResult.NewRow()
                            For i As Integer = 0 To row.Count - 1
                                nr.Item(i) = row(i)
                            Next
                            dtResult.Rows.Add(nr)
                        Next
                        _datatable.Merge(dtResult)
                        'NPOIXlsUtil.RenderDataTableToExcel(_datatable, "D:\GATrafficCampaign.xls")
                        'Console.WriteLine("merged {0} rows", dtResult.Rows.Count)
                    End If
                    ' Prepare the next page of results             
                    request.StartIndex = request.StartIndex + request.MaxResults
                    ' Execute and process the next page request
                    result = request.Execute()
                End While


                _datatable.Columns.Add("WEB_DOMAIN")
                _datatable.AcceptChanges()

                For Each r As DataRow In _datatable.Rows
                    'r.Item("ROW_ID") = Left(Replace(Guid.NewGuid().ToString(), "-", ""), 15)
                    If _datatable.Columns.Contains("ga:SearchKeyword") Then
                        r.Item("ga:SearchKeyword") = System.Web.HttpUtility.UrlDecode(r.Item("ga:SearchKeyword"))
                        If r.Item("ga:SearchKeyword").ToString.Length > 1000 Then r.Item("ga:SearchKeyword") = Left(r.Item("ga:SearchKeyword"), 1000)
                    End If

                    If _datatable.Columns.Contains("ga:dateHour") Then
                        'Console.WriteLine("{0}:{1}", r.Item("ROW_ID"), r.Item("ga:dateHour"))
                        'Threading.Thread.Sleep(100)
                        Dim RowTimeStamp As Date = Date.ParseExact(r.Item("ga:dateHour"), "yyyyMMddHH", New Globalization.CultureInfo("en-US"))
                        r.Item("ga:dateHour") = RowTimeStamp.ToString("yyyy/MM/dd HH:mm:ss")
                    End If
                    'r.Item("ga:landingPagePath") = WebDomain + r.Item("ga:landingPagePath").ToString()
                    Dim landingPagePath As String = r.Item("ga:landingPagePath")
                    If landingPagePath.Contains("/") AndAlso Not String.IsNullOrEmpty(landingPagePath.Substring(0, landingPagePath.IndexOf("/"))) Then
                        r.Item("WEB_DOMAIN") = landingPagePath.Substring(0, landingPagePath.IndexOf("/"))
                    End If
                    If r.Item("WEB_DOMAIN").ToString.Length > 100 Then r.Item("WEB_DOMAIN") = Left(r.Item("WEB_DOMAIN"), 100)
                    'Console.WriteLine("web domain:" + r.Item("WEB_DOMAIN"))
                    'Console.ReadKey()
                Next

                Dim curationConn As New SqlConnection(CurationPoolConnStr)
                curationConn.Open()
                Dim cmd As New SqlClient.SqlCommand( _
                       " delete from CurationPool.dbo." + writeToTable + " " + _
                       " where dateHour>='" + StartDate.ToString("yyyy-MM-dd") + "' and dateHour<='" + EndDate.ToString("yyyy-MM-dd") + "' " + _
                       " ", curationConn)
                cmd.CommandTimeout = 9999 : cmd.ExecuteNonQuery()
                Dim bk As New SqlBulkCopy(curationConn)
                bk.DestinationTableName = writeToTable
                bk.BulkCopyTimeout = 99999
                Try
                    bk.WriteToServer(_datatable)
                Catch ex As InvalidOperationException
                    Console.WriteLine(ex.ToString())
                    For Each col As DataColumn In _datatable.Columns
                        Dim maxLen As Integer = 0
                        For Each qr As DataRow In _datatable.Rows
                            If qr.Item(col.ColumnName) IsNot DBNull.Value Then
                                If qr.Item(col.ColumnName).ToString().Length > maxLen Then maxLen = qr.Item(col.ColumnName).ToString.Length
                            End If
                        Next
                        Console.WriteLine("{0}'s max len:{1};", col.ColumnName, maxLen)
                    Next
                    Console.Read()
                End Try

                Console.WriteLine("wrote {0} rows to " + writeToTable, _datatable.Rows.Count)
                curationConn.Close()
                FromDate = DateAdd(DateInterval.Day, intDayDiff + 1, StartDate)
            Next
            'End of gtype (traffic, search)
            'Exit For
        Next

        Console.WriteLine("done")
        'Console.Read()

    End Sub

    Sub GATrafficCampaign()
        Dim dirInfo As New IO.DirectoryInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim scopes As String() = New String() {AnalyticsService.Scope.AnalyticsReadonly}
        Dim keyFilePath As String = dirInfo.Parent.Parent.Parent.FullName + "\key.p12"
        Dim serviceAccountEmail As String = "127038164893-9qlqfak1plsd7uv9t5jmr9l12dkj86ub@developer.gserviceaccount.com"
        Dim certificate = New X509Certificate2(keyFilePath, "notasecret", X509KeyStorageFlags.Exportable)
        Dim Initializer1 As New ServiceAccountCredential.Initializer(serviceAccountEmail)
        Initializer1.Scopes = scopes : Initializer1.FromCertificate(certificate)
        Dim credential As New ServiceAccountCredential(Initializer1)
        Dim Initializer2 As New BaseClientService.Initializer()
        Initializer2.HttpClientInitializer = credential : Initializer2.ApplicationName = "API Projects"
        Dim service As New AnalyticsService(Initializer2)
        'End of Init

        Dim writeToTable = "GOOGLE_GA_TRAFFIC_SEM", intDayDiff = 6
        Dim FromDate As Date = DateAdd(DateInterval.Day, -10, Now)
        'FromDate = New Date(2015, 12, 3)
        Dim iteration As Integer = DateDiff(DateInterval.Day, FromDate, DateAdd(DateInterval.Day, -1, Now)) / intDayDiff
        Dim Metrics As String = "ga:pageviews,ga:bounces"
        Dim Dimension As String = "ga:dateHour,ga:country,ga:city,ga:landingPagePath,ga:medium,ga:source,ga:campaign"

        For idxDays As Integer = 0 To iteration
            Dim StartDate As New Date(FromDate.Year, FromDate.Month, FromDate.Day), EndDate As Date = DateAdd(DateInterval.Day, intDayDiff, StartDate)
            If DateDiff(DateInterval.Day, Now, StartDate) >= 0 Then Exit For
            Console.WriteLine("{0}~{1} of {2}", StartDate.ToString("yyyy-MM-dd"), EndDate.ToString("yyyy-MM-dd"), "Traffic Source")
            Dim _datatable As New DataTable
            Dim request As DataResource.GaResource.GetRequest = service.Data.Ga.Get("ga:9115484", StartDate.ToString("yyyy-MM-dd"), EndDate.ToString("yyyy-MM-dd"), Metrics)
            request.Dimensions = Dimension : request.MaxResults = 10000 : request.StartIndex = 1
            Dim result As GaData = request.Execute()

            While result.Rows IsNot Nothing
                If result.NextLink Is Nothing Then
                    If request.MaxResults <= 1 Then
                        Exit While
                    Else
                        If request.MaxResults > 100 Then
                            request.MaxResults /= 10
                        Else
                            request.MaxResults /= 2
                        End If
                    End If
                Else
                    Dim dtResult As New DataTable
                    For Each headers In result.ColumnHeaders
                        dtResult.Columns.Add(headers.Name)
                    Next

                    For Each row In result.Rows
                        Dim nr As DataRow = dtResult.NewRow()
                        For i As Integer = 0 To row.Count - 1
                            nr.Item(i) = row(i)
                        Next
                        dtResult.Rows.Add(nr)
                    Next
                    _datatable.Merge(dtResult)
                    If dtResult.Rows.Count = 1 Then Exit While
                    'NPOIXlsUtil.RenderDataTableToExcel(_datatable, "D:\GATraffic_SourceMediumOrganic.xls")
                    Console.WriteLine("merged {0} rows", dtResult.Rows.Count)
                End If
                ' Prepare the next page of results             
                request.StartIndex = request.StartIndex + request.MaxResults
                ' Execute and process the next page request
                result = request.Execute()
            End While
            _datatable.Columns.Add("WEB_DOMAIN")
            _datatable.AcceptChanges()

            For Each r As DataRow In _datatable.Rows

                If _datatable.Columns.Contains("ga:dateHour") Then
                    'Console.WriteLine("{0}:{1}", r.Item("ROW_ID"), r.Item("ga:dateHour"))
                    'Threading.Thread.Sleep(100)
                    Dim RowTimeStamp As Date = Date.ParseExact(r.Item("ga:dateHour"), "yyyyMMddHH", New Globalization.CultureInfo("en-US"))
                    r.Item("ga:dateHour") = RowTimeStamp.ToString("yyyy/MM/dd HH:mm:ss")
                End If
                'r.Item("ga:landingPagePath") = WebDomain + r.Item("ga:landingPagePath").ToString()
                If r.Item("ga:landingPagePath").ToString.Length > 1500 Then r.Item("ga:landingPagePath") = Left(r.Item("ga:landingPagePath"), 1500)
                If r.Item("ga:source").ToString.Length > 200 Then r.Item("ga:source") = Left(r.Item("ga:source"), 200)
                Dim landingPagePath As String = r.Item("ga:landingPagePath")
                If landingPagePath.Contains("/") AndAlso Not String.IsNullOrEmpty(landingPagePath.Substring(0, landingPagePath.IndexOf("/"))) Then
                    r.Item("WEB_DOMAIN") = landingPagePath.Substring(0, landingPagePath.IndexOf("/"))
                End If
                If r.Item("WEB_DOMAIN").ToString.Length > 100 Then r.Item("WEB_DOMAIN") = Left(r.Item("WEB_DOMAIN"), 100)
                'Console.WriteLine("web domain:" + r.Item("WEB_DOMAIN"))
                'Console.ReadKey()
            Next

            Dim myLocalConn As New SqlConnection(MyLocal)
            myLocalConn.Open()
            Dim cmd As New SqlClient.SqlCommand( _
                   " delete from MyLocal.dbo." + writeToTable + " " + _
                   " where dateHour>='" + StartDate.ToString("yyyy-MM-dd 00:00:00") + "' and dateHour<='" + EndDate.ToString("yyyy-MM-dd 23:59:59") + "' " + _
                   " ", myLocalConn)
            cmd.CommandTimeout = 9999 : cmd.ExecuteNonQuery()
            Dim bk As New SqlBulkCopy(myLocalConn)
            bk.DestinationTableName = writeToTable
            bk.BulkCopyTimeout = 99999
            Try
                bk.WriteToServer(_datatable)
            Catch ex As InvalidOperationException
                Console.WriteLine(ex.ToString())
                For Each col As DataColumn In _datatable.Columns
                    Dim maxLen As Integer = 0
                    For Each qr As DataRow In _datatable.Rows
                        If qr.Item(col.ColumnName) IsNot DBNull.Value Then
                            If qr.Item(col.ColumnName).ToString().Length > maxLen Then maxLen = qr.Item(col.ColumnName).ToString.Length
                        End If
                    Next
                    Console.WriteLine("{0}'s max len:{1};", col.ColumnName, maxLen)
                Next
                Console.Read()
            End Try

            Console.WriteLine("wrote {0} rows to " + writeToTable, _datatable.Rows.Count)
            myLocalConn.Close()
            FromDate = DateAdd(DateInterval.Day, intDayDiff + 1, StartDate)
        Next
    End Sub

    Sub GASearchSourceMedium()
        'Init API login info
        Dim dirInfo As New IO.DirectoryInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim scopes As String() = New String() {AnalyticsService.Scope.AnalyticsReadonly}
        Dim keyFilePath As String = dirInfo.Parent.Parent.Parent.FullName + "\key.p12"
        Dim serviceAccountEmail As String = "127038164893-9qlqfak1plsd7uv9t5jmr9l12dkj86ub@developer.gserviceaccount.com"
        Dim certificate = New X509Certificate2(keyFilePath, "notasecret", X509KeyStorageFlags.Exportable)
        Dim Initializer1 As New ServiceAccountCredential.Initializer(serviceAccountEmail)
        Initializer1.Scopes = scopes : Initializer1.FromCertificate(certificate)
        Dim credential As New ServiceAccountCredential(Initializer1)
        Dim Initializer2 As New BaseClientService.Initializer()
        Initializer2.HttpClientInitializer = credential : Initializer2.ApplicationName = "API Projects"
        Dim service As New AnalyticsService(Initializer2)
        'End of Init

        Dim writeToTable = "GOOGLE_GA_SEARCH_SEM", intDayDiff = 6
        Dim FromDate As Date = DateAdd(DateInterval.Day, -10, Now)
        'FromDate = New Date(2015, 9, 20)
        Dim iteration As Integer = DateDiff(DateInterval.Day, FromDate, DateAdd(DateInterval.Day, -1, Now)) / intDayDiff

        Dim Metrics As String = "ga:entrances,ga:pageviews"
        Dim Dimension As String = "ga:searchKeyword,ga:dateHour,ga:country,ga:city,ga:landingPagePath,ga:medium,ga:source"

        For idxDays As Integer = 0 To iteration
            Dim StartDate As New Date(FromDate.Year, FromDate.Month, FromDate.Day), EndDate As Date = DateAdd(DateInterval.Day, intDayDiff, StartDate)
            If DateDiff(DateInterval.Day, Now, StartDate) >= 0 Then Exit For
            Console.WriteLine("{0}~{1} of {2}", StartDate.ToString("yyyy-MM-dd"), EndDate.ToString("yyyy-MM-dd"), "Search Medium")
            Dim _datatable As New DataTable
            Dim request As DataResource.GaResource.GetRequest = service.Data.Ga.Get("ga:9115484", StartDate.ToString("yyyy-MM-dd"), EndDate.ToString("yyyy-MM-dd"), Metrics)
            request.Dimensions = Dimension : request.MaxResults = 10000 : request.StartIndex = 1
            Dim result As GaData = request.Execute()
          
            While result.Rows IsNot Nothing
                If result.NextLink Is Nothing Then
                    If request.MaxResults <= 1 Then
                        Exit While
                    Else
                        If request.MaxResults > 100 Then
                            request.MaxResults /= 10
                        Else
                            request.MaxResults /= 2
                        End If
                    End If
                Else
                    Dim dtResult As New DataTable
                    For Each headers In result.ColumnHeaders
                        dtResult.Columns.Add(headers.Name)
                    Next

                    For Each row In result.Rows
                        Dim nr As DataRow = dtResult.NewRow()
                        For i As Integer = 0 To row.Count - 1
                            nr.Item(i) = row(i)
                        Next
                        dtResult.Rows.Add(nr)
                    Next
                    _datatable.Merge(dtResult)
                    'NPOIXlsUtil.RenderDataTableToExcel(_datatable, "D:\GASearchOrganic.xls")
                    'Console.WriteLine("merged {0} rows", dtResult.Rows.Count)
                End If
                ' Prepare the next page of results             
                request.StartIndex = request.StartIndex + request.MaxResults
                ' Execute and process the next page request
                result = request.Execute()
            End While
            _datatable.Columns.Add("WEB_DOMAIN")
            _datatable.AcceptChanges()

            For Each r As DataRow In _datatable.Rows
               
                If _datatable.Columns.Contains("ga:SearchKeyword") Then
                    r.Item("ga:SearchKeyword") = System.Web.HttpUtility.UrlDecode(r.Item("ga:SearchKeyword"))
                    If r.Item("ga:SearchKeyword").ToString.Length > 1000 Then r.Item("ga:SearchKeyword") = Left(r.Item("ga:SearchKeyword"), 1000)
                End If

                If _datatable.Columns.Contains("ga:dateHour") Then
                    'Console.WriteLine("{0}:{1}", r.Item("ROW_ID"), r.Item("ga:dateHour"))
                    'Threading.Thread.Sleep(100)
                    Dim RowTimeStamp As Date = Date.ParseExact(r.Item("ga:dateHour"), "yyyyMMddHH", New Globalization.CultureInfo("en-US"))
                    r.Item("ga:dateHour") = RowTimeStamp.ToString("yyyy/MM/dd HH:mm:ss")
                End If

                If r.Item("ga:landingPagePath").ToString.Length > 1500 Then r.Item("ga:landingPagePath") = Left(r.Item("ga:landingPagePath"), 1500)

                'r.Item("ga:landingPagePath") = WebDomain + r.Item("ga:landingPagePath").ToString()
                If r.Item("ga:source").ToString.Length > 200 Then r.Item("ga:source") = Left(r.Item("ga:source"), 200)
                Dim landingPagePath As String = r.Item("ga:landingPagePath")
                If landingPagePath.Contains("/") AndAlso Not String.IsNullOrEmpty(landingPagePath.Substring(0, landingPagePath.IndexOf("/"))) Then
                    r.Item("WEB_DOMAIN") = landingPagePath.Substring(0, landingPagePath.IndexOf("/"))
                End If
                If r.Item("WEB_DOMAIN").ToString.Length > 100 Then r.Item("WEB_DOMAIN") = Left(r.Item("WEB_DOMAIN"), 100)
                'Console.WriteLine("web domain:" + r.Item("WEB_DOMAIN"))
                'Console.ReadKey()
            Next

            Dim myLocalConn As New SqlConnection(MyLocal)
            myLocalConn.Open()
            Dim cmd As New SqlClient.SqlCommand( _
                   " delete from MyLocal.dbo." + writeToTable + " " + _
                   " where dateHour>='" + StartDate.ToString("yyyy-MM-dd 00:00:00") + "' and dateHour<='" + EndDate.ToString("yyyy-MM-dd 23:59:59") + "' " + _
                   " ", myLocalConn)
            cmd.CommandTimeout = 9999 : cmd.ExecuteNonQuery()
            Dim bk As New SqlBulkCopy(myLocalConn)
            bk.DestinationTableName = writeToTable
            bk.BulkCopyTimeout = 99999
            Try
                bk.WriteToServer(_datatable)
            Catch ex As InvalidOperationException
                Console.WriteLine(ex.ToString())
                For Each col As DataColumn In _datatable.Columns
                    Dim maxLen As Integer = 0
                    For Each qr As DataRow In _datatable.Rows
                        If qr.Item(col.ColumnName) IsNot DBNull.Value Then
                            If qr.Item(col.ColumnName).ToString().Length > maxLen Then maxLen = qr.Item(col.ColumnName).ToString.Length
                        End If
                    Next
                    Console.WriteLine("{0}'s max len:{1};", col.ColumnName, maxLen)
                Next
                Console.Read()
            End Try

            Console.WriteLine("wrote {0} rows to " + writeToTable, _datatable.Rows.Count)
            myLocalConn.Close()
            FromDate = DateAdd(DateInterval.Day, intDayDiff + 1, StartDate)
        Next
    End Sub

#Region "Only kept for Reference"

    Sub GetAccountSummaries(service As AnalyticsService)
        Dim list As ManagementResource.AccountSummariesResource.ListRequest = service.Management.AccountSummaries.List()
        list.MaxResults = 1000
        ' Maximum number of Account Summaries to return per request. 
        Dim feed As AccountSummaries = list.Execute()
        Dim allRows As New List(Of AccountSummary)

        '''/ Loop through until we arrive at an empty page
        While feed.Items IsNot Nothing
            allRows.AddRange(feed.Items)

            ' We will know we are on the last page when the next page token is
            ' null.
            ' If this is the case, break.
            If feed.NextLink Is Nothing Then
                Exit While
            End If

            ' Prepare the next page of results             
            list.StartIndex = feed.StartIndex + list.MaxResults
            ' Execute and process the next page request

            feed = list.Execute()
        End While

        feed.Items = allRows
        ' feed.Items not contains all of the rows even if there are more then 1000 


        For Each account As AccountSummary In feed.Items
            ' Account
            Console.WriteLine("Account: " + account.Name + "(" + account.Id + ")")
            For Each wp As WebPropertySummary In account.WebProperties
                ' Web Properties within that account
                Console.WriteLine(vbTab & "Web Property: " + wp.Name + "(" + wp.Id + ")")

                'Don't forget to check its not null. Believe it or not it could be.
                If wp.Profiles IsNot Nothing Then
                    For Each profile As ProfileSummary In wp.Profiles
                        ' Profiles with in that web property.
                        Console.WriteLine(vbTab & vbTab & "Profile: " + profile.Name + "(" + profile.Id + ")")
                    Next
                End If
            Next
        Next
    End Sub

    Sub GetAccounts(service As AnalyticsService)
        Dim list As ManagementResource.AccountsResource.ListRequest = service.Management.Accounts.List()
        list.MaxResults = 1000
        ' Maximum number of Accounts to return, per request. 
        Dim feed As Accounts = list.Execute()

        For Each account As Account In feed.Items
            ' Account
            Console.WriteLine(String.Format("Account: {0}({1})", account.Name, account.Id))
        Next
    End Sub

    Sub GetProfiles(service As AnalyticsService, accountId As String, WebPropertyId As String)
        'account id: 2826869
        'web property id: UA-2826869-26
        Dim list As ManagementResource.ProfilesResource.ListRequest = service.Management.Profiles.List(accountId, WebPropertyId)
        Dim feed As Profiles = list.Execute()

        For Each profile As Profile In feed.Items
            Console.WriteLine(String.Format(vbTab & vbTab & "Profile: {0}({1})", profile.Name, profile.Id))
        Next


    End Sub

    Sub GetAccountWebProperties(service As AnalyticsService, accountId As String)
        Dim list As ManagementResource.WebpropertiesResource.ListRequest = service.Management.Webproperties.List(accountId)
        Dim webProperties As Webproperties = list.Execute()

        For Each wp As Webproperty In webProperties.Items
            ' Web Properties within that account
            Console.WriteLine(String.Format(vbTab & "Web Property: {0}({1})", wp.Name, wp.Id))
        Next

    End Sub

#End Region
End Module
