Imports OpenQA.Selenium
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.Support.UI

Module Module1
    'Public ID As String = "rex.chou" : Public PWD As String = "braceup4$" : Public LiveChatID As String = "68676965", rnd As Random
    'Public ID As String = "tony.fu" : Public PWD As String = "tonyfu7324" : Public LiveChatID As String = "68676965", rnd As Random
    Public ID As String = "rudy.wang" : Public PWD As String = "advantech3" : Public LiveChatID As String = "68676965", rnd As Random '2018/04/16
    Public recall As Integer = 0
    Sub Main()
        StartProcedure()
    End Sub

    Sub StartProcedure()
        If recall >= 5 Then
            CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Import Livechat Failed", "Time out", True, "", "")
            Exit Sub
        Else
            'Selenium Firefox temp file path C:\Documents and Settings\ebiz.aeu\Local Settings\Temp\2
            'Win 2012 temp path: C:\Users\tc.chen\AppData\Local\Temp\2

            CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Start Import Livechat", Now.ToString, True, "", "")

            rnd = New Random()
            Dim driver = LoginLivePerson()

            Dim AMKTWS As New AOLWS.IntApi
            AMKTWS.Timeout = 1000 * 60 * 90

            Try
                Threading.Thread.Sleep(rnd.Next(9999, 60000))
                'Click "Reporting & Analytics" to expand function links
                driver.FindElement(By.CssSelector(".navigationTable > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(7) > td:nth-child(1) > div:nth-child(1) > a:nth-child(1) > img:nth-child(1)")).Click()

                Threading.Thread.Sleep(500)
                'Click "Transcripts"
                driver.FindElement(By.CssSelector("#innerTable7 > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > a:nth-child(1)")).Click()
                Console.WriteLine("Entering Transcript page")
                Threading.Thread.Sleep(rnd.Next(1234, 3210))


                Dim _date As Date = DateAdd(DateInterval.Day, -1, Now)
                Dim StartDateTime As New Date(Year(_date), Month(_date), Day(_date), 0, 0, 0), EndDateTime As New Date(Year(_date), Month(_date), Day(_date), 23, 59, 59)
                'Dim StartDateTime As New Date(2017, 3, 10, 0, 0, 0), EndDateTime As New Date(2017, 3, 10, 23, 59, 59)
                Dim SessionTypes() As SessionType = {SessionType.Chats, SessionType.Calls}
                For Each sType In SessionTypes
                    Dim LiveChatDs As LivePersonDataSet = GetChatsAndCallsWithinDateRange(driver, StartDateTime, EndDateTime, sType)

                    Dim SessionGenInfo = From s In LiveChatDs.Sessions Join g In LiveChatDs.GenInfoRec On s.RealTimeSessionID Equals g.RealTimeSessionID
                                         Select s.RealTimeSessionID, s.Start, s.Duration, s.Operator, s.skill, s.phoneNumber, s.email, g.Country, s.identifier, g.Org

                    Dim ImportChatContacts As New List(Of AOLWS.LiveChatContact), ImportChatContactURLs As New List(Of AOLWS.LiveChatURL), ImportErrMsg As String = ""

                    For Each SGInfo In SessionGenInfo
                        Dim LCContact1 As New AOLWS.LiveChatContact()
                        With LCContact1
                            .Country = SGInfo.Country : .Duration = SGInfo.Duration : .email = SGInfo.email
                            .identifier = SGInfo.identifier : .Operator = SGInfo.Operator : .Org = SGInfo.Org
                            .phoneNumber = SGInfo.phoneNumber : .RealTimeSessionID = SGInfo.RealTimeSessionID
                            .skill = SGInfo.skill : .Start = CDate(SGInfo.Start)
                            'TC added this intentionally to make Olivia's chat looks like abandoned
                            .Operator = ""
                        End With
                        ImportChatContacts.Add(LCContact1)

                        Dim naviRecords = From q In LiveChatDs.Navigations Where q.RealTimeSessionID = LCContact1.RealTimeSessionID Order By q.SeqNo

                        If naviRecords.Count > 0 Then

                            Dim naviStartDate As DateTime = DateAdd(DateInterval.Minute, naviRecords.Count * -1, CDate(LCContact1.Start))

                            For recCount As Integer = 0 To naviRecords.Count - 1
                                Dim LCURL1 As New AOLWS.LiveChatURL()
                                With LCURL1
                                    .RealTimeSessionID = naviRecords(recCount).RealTimeSessionID : .SeqNo = naviRecords(recCount).SeqNo : .URL = naviRecords(recCount).URL
                                    .Timestamp = DateAdd(DateInterval.Minute, recCount, naviStartDate)
                                End With
                                ImportChatContactURLs.Add(LCURL1)
                            Next

                        End If
                    Next

                    Dim ds As New DataSet

                    Dim dtChatSessions As DataTable = Helper.ListToDataTable(LiveChatDs.Sessions)

                    For Each r As DataRow In dtChatSessions.Rows
                        Dim TS() As String = Split(r.Item("Start").ToString(), vbCrLf)
                        r.Item("Start") = Date.ParseExact(TS(1), "MM/dd/yyyy", New Globalization.CultureInfo("en-US")).ToString("yyyy/MM/dd") + " " + TS(0)
                    Next

                    ds.Tables.Add(dtChatSessions) : ds.Tables.Add(Helper.ListToDataTable(LiveChatDs.GenInfoRec))
                    ds.Tables.Add(Helper.ListToDataTable(LiveChatDs.Transcripts)) : ds.Tables.Add(Helper.ListToDataTable(LiveChatDs.Navigations))

                    'SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Start Import Livechat Dataset", "1: " + ds.Tables(0).Rows.Count.ToString + "<br/>2: " + ds.Tables(1).Rows.Count.ToString + "<br/>3: " + ds.Tables(2).Rows.Count.ToString + "<br/>4: " + ds.Tables(3).Rows.Count.ToString, True, "", "")

                    'NPOIXlsUtil.RenderDataSetToExcel(ds, _
                    'String.Format("LivePerson_{0}_Sessions_{1}-{2}_SavedOn_{3}.xls", sType.ToString(), StartDateTime.ToString("yyyyMMddHHmm"), EndDateTime.ToString("yyyyMMddHHmm"), Now.ToString("yyyyMMddHHmm")))

                    AMKTWS.ImportLiveChatByDataset(ds, String.Format("LivePerson_{0}_Sessions_{1}-{2}_SavedOn_{3}", sType.ToString(), StartDateTime.ToString("yyyyMMddHHmm"), EndDateTime.ToString("yyyyMMddHHmm"), Now.ToString("yyyyMMddHHmm")), "")
                    'AMKTWS.ImportLiveChat(ImportChatContacts.ToArray(), ImportChatContactURLs.ToArray(), ImportErrMsg)

                    'SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "End Import Livechat Dataset", Now.ToString, True, "", "")

                    'Go back to query page
                    driver.FindElement(By.CssSelector("body > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(2) > table > tbody > tr:nth-child(7) > td > table > tbody > tr > td.control > form > input")).Click()

                Next
            Catch ex As Exception
                CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Import Livechat Failed", ex.ToString, True, "", "")
            End Try


            CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "End Import Livechat", Now.ToString, True, "", "")


            'Console.WriteLine("Session table is not ready, press any key to exit...") : Console.ReadKey()
            'Logout
            driver.FindElement(By.CssSelector("td.printHidden > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(3) > a:nth-child(1) > img:nth-child(1)")).Click()
            Threading.Thread.Sleep(rnd.Next(1234, 3210))
            driver.Close()

            'SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Start Import Livechat Excel", Now.ToString, True, "", "")

            '20150818 Rudy: Call WS to Import Live Chat Contact, URL, and Lead
            'AMKTWS.ImportLiveChatByExcel("")

            'SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "End Import Livechat Excel", Now.ToString, True, "", "")
            'Console.WriteLine("Logged out, press any key to exit...") : Console.ReadLine()
        End If
    End Sub

    Sub GetOldData()
        'Dim SelectedSessionType As SessionType = SessionType.Chats
        'Dim FromDate As Date = Date.MinValue, ToDate As Date = New Date(2015, 5, 14)
        'For i As Integer = 1 To 2
        '    FromDate = DateAdd(DateInterval.Day, 0, ToDate)
        '    Console.WriteLine("Search {0} to {1}", FromDate.ToString("yyyyMMdd"), ToDate.ToString("yyyyMMdd"))
        '    'Start Date
        '    '#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > input:nth-child(1)
        '    driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > input:nth-child(1)")).Clear()
        '    driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > input:nth-child(1)")).SendKeys(FromDate.ToString("M/d/yyyy"))
        '    'End Date
        '    driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2) > input:nth-child(1)")).Clear()
        '    driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2) > input:nth-child(1)")).SendKeys(ToDate.ToString("M/d/yyyy"))

        '    'End Time Hour
        '    Dim dlEndHour = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(5) > select:nth-child(1)")))
        '    dlEndHour.SelectByText("23")
        '    'End Time Min
        '    Dim dlEndMin = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(5) > select:nth-child(3)")))
        '    dlEndMin.SelectByText("59")
        '    'Session Type
        '    Dim dlSessionType = New SelectElement(driver.FindElement(By.CssSelector("tr.tabledata:nth-child(1) > td:nth-child(2) > select:nth-child(1)")))
        '    dlSessionType.SelectByText(SelectedSessionType.ToString())
        '    'Operator Type
        '    Dim dlOperator = New SelectElement(driver.FindElement(By.CssSelector("#opId")))
        '    dlOperator.SelectByText("All Sessions")

        '    'Timezone
        '    Dim dlTimezone = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(2) > select:nth-child(1)")))
        '    dlTimezone.SelectByValue("Asia/Taipei")
        '    'Result per page
        '    Dim dlResultPerPage = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table > tbody > tr:nth-child(7) > td > table > tbody > tr:nth-child(15) > td:nth-child(2) > select")))
        '    dlResultPerPage.SelectByText("100") ' 100 results per page

        '    'List Session Button
        '    driver.FindElement(By.CssSelector("td.control:nth-child(1) > input:nth-child(1)")).Click()
        '    Threading.Thread.Sleep(2266)

        '    Dim ChatSessions As New List(Of Session), ChatGenInfoList As New List(Of GeneralChatInfo), ChatTranscriptList As New List(Of ChatTranscript)
        '    Dim NavigationList As New List(Of Navigation)

        '    If WebDriverUtil.IsElementReady(driver, By.CssSelector("#transcriptTable"), 60) Then
        '        'Try
        '        While True
        '            Dim resultTable = driver.FindElement(By.CssSelector("#transcriptTable"))
        '            Dim TableDataRows = resultTable.FindElements(By.ClassName("tabledata"))
        '            For Each dataRow In TableDataRows
        '                Dim TDs = dataRow.FindElements(By.TagName("td")), surveyTDs = dataRow.FindElements(By.ClassName("GenTD04"))
        '                Dim session1 As New Session
        '                With session1
        '                    .Start = TDs(0).Text : .Duration = TDs(1).Text : .Operator = TDs(2).Text : .RealTimeSessionID = TDs(3).Text : .message = TDs(4).Text
        '                    .Extension = TDs(5).Text : .Ticket = TDs(6).Text : .skill = TDs(7).Text : .phoneNumber = TDs(8).Text : .email = TDs(9).Text
        '                    .subject = TDs(10).Text : .Country = TDs(11).Text : .identifier = TDs(12).Text
        '                    If surveyTDs IsNot Nothing AndAlso surveyTDs.Count >= 4 Then
        '                        .Survey_PreChat = IIf(surveyTDs(0).FindElements(By.TagName("img")).Count > 0, "y", "n")
        '                        .Survey_Exit = IIf(surveyTDs(1).FindElements(By.TagName("img")).Count > 0, "y", "n")
        '                        .Survey_Operator = IIf(surveyTDs(2).FindElements(By.TagName("img")).Count > 0, "y", "n")
        '                        .Survey_Offline = IIf(surveyTDs(3).FindElements(By.TagName("img")).Count > 0, "y", "n")
        '                    End If

        '                    'Click into Chat Detail and get conversation log and navigated URLs
        '                    Dim act As New Interactions.Actions(driver)
        '                    act.KeyDown(Keys.Shift).Click(TDs(0).FindElement(By.ClassName("chats"))).KeyUp(Keys.Shift).Build().Perform()
        '                    Threading.Thread.Sleep(rnd.Next(333, 1501))

        '                    driver.SwitchTo().Window(driver.WindowHandles(1))
        '                    Dim ChatDetails1 As ChatDetails = GetChatDetail(driver)
        '                    ChatGenInfoList.Add(ChatDetails1.GenInfo)
        '                    For Each ct In ChatDetails1.ChatTranscripts
        '                        ChatTranscriptList.Add(ct)
        '                    Next

        '                    For Each navi In ChatDetails1.Navigations
        '                        NavigationList.Add(navi)
        '                    Next

        '                    Threading.Thread.Sleep(rnd.Next(333, 1234))
        '                    driver.SwitchTo().Window(driver.WindowHandles(1)).Close()
        '                    driver.SwitchTo().Window(driver.WindowHandles(0))


        '                    'If ChatSessions.Count >= 10 AndAlso ChatSessions.Count Mod 10 = 0 Then
        '                    Console.WriteLine("Stored {0} sessions", ChatSessions.Count)
        '                    Console.WriteLine("Adding session: Start:{0}, Duration:{1}, Operator:{2}", .Start, .Duration, .Operator)
        '                    'End If

        '                End With
        '                ChatSessions.Add(session1)
        '                'If ChatGenInfoList.Count >= 2 AndAlso ChatTranscriptList.Count >= 1 Then Exit While
        '            Next

        '            Threading.Thread.Sleep(rnd.Next(333, 2345))
        '            'click next results link
        '            If WebDriverUtil.IsElementReady(driver, By.CssSelector("tr.copysmall > td:nth-child(2) > a:nth-child(1)"), 10) Then
        '                Console.WriteLine("Next results")
        '                driver.FindElement(By.CssSelector("tr.copysmall > td:nth-child(2) > a:nth-child(1)")).Click()
        '                Threading.Thread.Sleep(rnd.Next(123, 1234))
        '            Else
        '                Console.WriteLine("No more results") : Exit While
        '            End If
        '        End While
        '        Dim ds As New DataSet

        '        Dim dtChatSessions As DataTable = Helper.ListToDataTable(ChatSessions)

        '        For Each r As DataRow In dtChatSessions.Rows
        '            Dim TS() As String = Split(r.Item("Start").ToString(), vbCrLf)
        '            r.Item("Start") = Date.ParseExact(TS(1), "MM/dd/yyyy", New Globalization.CultureInfo("en-US")).ToString("yyyy/MM/dd") + " " + TS(0)
        '        Next

        '        ds.Tables.Add(dtChatSessions) : ds.Tables.Add(Helper.ListToDataTable(ChatGenInfoList))
        '        ds.Tables.Add(Helper.ListToDataTable(ChatTranscriptList)) : ds.Tables.Add(Helper.ListToDataTable(NavigationList))

        '        NPOIXlsUtil.RenderDataSetToExcel(ds, _
        '        String.Format("LivePerson_{0}_Sessions_{1}-{2}_SavedOn_{3}.xls", SelectedSessionType.ToString(), FromDate.ToString("yyyyMMdd"), ToDate.ToString("yyyyMMdd"), Now.ToString("yyyyMMddHHmm")))

        '    Else
        '        'Exit For
        '    End If

        '    ToDate = DateAdd(DateInterval.Day, 1, FromDate)
        '    'Go back to query page
        '    driver.FindElement(By.CssSelector("body > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(2) > table > tbody > tr:nth-child(7) > td > table > tbody > tr > td.control > form > input")).Click()
        '    Console.WriteLine("Waiting for several minutes then perform next query...")
        '    Threading.Thread.Sleep(rnd.Next(1 * 60 * 1000, 3 * 60 * 1000))

        'Next

    End Sub

    Function GetChatsAndCallsWithinDateRange(driver As FirefoxDriver, FromDate As Date, ToDate As Date, SelectedSessionType As SessionType) As LivePersonDataSet
        Try
            Dim LivePersonDataSet1 As New LivePersonDataSet()
            Console.WriteLine("Search {0} to {1}, type {2}", FromDate.ToString("yyyyMMddHHmm"), ToDate.ToString("yyyyMMddHHmm"), SelectedSessionType.ToString())
            'Start Date
            driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > input:nth-child(1)")).Clear()
            driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > input:nth-child(1)")).SendKeys(FromDate.ToString("M/d/yyyy"))
            'End Date
            driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2) > input:nth-child(1)")).Clear()
            driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2) > input:nth-child(1)")).SendKeys(ToDate.ToString("M/d/yyyy"))

            'Begin Time Hour
            Dim dlBeginHour = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(5) > select:nth-child(1)")))
            dlBeginHour.SelectByText(FromDate.ToString("HH"))
            'Begin Time Min
            Dim dlBeginMin = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(5) > select:nth-child(3)")))
            dlBeginMin.SelectByText(FromDate.ToString("mm"))

            'End Time Hour
            Dim dlEndHour = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(5) > select:nth-child(1)")))
            dlEndHour.SelectByText(ToDate.ToString("HH"))
            'End Time Min
            Dim dlEndMin = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(5) > select:nth-child(3)")))
            dlEndMin.SelectByText(ToDate.ToString("mm"))
            'Session Type
            Dim dlSessionType = New SelectElement(driver.FindElement(By.CssSelector("tr.tabledata:nth-child(1) > td:nth-child(2) > select:nth-child(1)")))
            dlSessionType.SelectByText(SelectedSessionType.ToString())
            'Operator Type
            Dim dlOperator = New SelectElement(driver.FindElement(By.CssSelector("#opId")))
            dlOperator.SelectByText("All Sessions")

            'Timezone
            Dim dlTimezone = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(2) > select:nth-child(1)")))
            dlTimezone.SelectByValue("Asia/Taipei")
            'Result per page
            Dim dlResultPerPage = New SelectElement(driver.FindElement(By.CssSelector("#byDateDiv > table > tbody > tr:nth-child(7) > td > table > tbody > tr:nth-child(15) > td:nth-child(2) > select")))
            dlResultPerPage.SelectByText("100") ' 100 results per page

            'List Session Button
            driver.FindElement(By.CssSelector("td.control:nth-child(1) > input:nth-child(1)")).Click()
            Threading.Thread.Sleep(2266)

            Dim ChatSessions As New List(Of Session), ChatGenInfoList As New List(Of GeneralChatInfo), ChatTranscriptList As New List(Of ChatTranscript)
            Dim NavigationList As New List(Of Navigation)

            If WebDriverUtil.IsElementReady(driver, By.CssSelector("#transcriptTable"), 120) Then
                'Try
                While True
                    Dim resultTable = driver.FindElement(By.CssSelector("#transcriptTable"))
                    Dim TableDataRows = resultTable.FindElements(By.ClassName("tabledata"))
                    For Each dataRow In TableDataRows
                        Dim TDs = dataRow.FindElements(By.TagName("td")), surveyTDs = dataRow.FindElements(By.ClassName("GenTD04"))
                        Dim session1 As New Session
                        With session1
                            .Start = TDs(0).Text : .Duration = TDs(1).Text : .Operator = TDs(2).Text : .RealTimeSessionID = TDs(3).Text : .message = TDs(4).Text
                            .Extension = TDs(5).Text : .Ticket = TDs(6).Text : .skill = TDs(7).Text : .phoneNumber = TDs(8).Text : .email = TDs(9).Text
                            .subject = TDs(10).Text : .Country = TDs(11).Text : .identifier = TDs(12).Text
                            If surveyTDs IsNot Nothing AndAlso surveyTDs.Count >= 4 Then
                                .Survey_PreChat = IIf(surveyTDs(0).FindElements(By.TagName("img")).Count > 0, "y", "n")
                                .Survey_Exit = IIf(surveyTDs(1).FindElements(By.TagName("img")).Count > 0, "y", "n")
                                .Survey_Operator = IIf(surveyTDs(2).FindElements(By.TagName("img")).Count > 0, "y", "n")
                                .Survey_Offline = IIf(surveyTDs(3).FindElements(By.TagName("img")).Count > 0, "y", "n")
                            End If

                            'Click into Chat Detail and get conversation log and navigated URLs
                            Dim act As New Interactions.Actions(driver)
                            act.KeyDown(Keys.Shift).Click(TDs(0).FindElement(By.ClassName("chats"))).KeyUp(Keys.Shift).Build().Perform()
                            Threading.Thread.Sleep(rnd.Next(333, 1501))

                            driver.SwitchTo().Window(driver.WindowHandles(1))
                            Dim ChatDetails1 As ChatDetails = GetChatDetail(driver)
                            ChatGenInfoList.Add(ChatDetails1.GenInfo)
                            For Each ct In ChatDetails1.ChatTranscripts
                                ChatTranscriptList.Add(ct)
                            Next

                            For Each navi In ChatDetails1.Navigations
                                NavigationList.Add(navi)
                            Next

                            Threading.Thread.Sleep(rnd.Next(333, 1234))
                            driver.SwitchTo().Window(driver.WindowHandles(1)).Close()
                            driver.SwitchTo().Window(driver.WindowHandles(0))


                            'If ChatSessions.Count >= 10 AndAlso ChatSessions.Count Mod 10 = 0 Then
                            Console.WriteLine("Stored {0} sessions", ChatSessions.Count)
                            Console.WriteLine("Adding session: Start:{0}, Duration:{1}, Operator:{2}", .Start, .Duration, .Operator)
                            'End If

                        End With
                        ChatSessions.Add(session1)
                        'If ChatGenInfoList.Count >= 2 AndAlso ChatTranscriptList.Count >= 1 Then Exit While
                    Next

                    Threading.Thread.Sleep(rnd.Next(333, 2345))
                    'click next results link
                    If WebDriverUtil.IsElementReady(driver, By.CssSelector("tr.copysmall > td:nth-child(2) > a:nth-child(1)"), 10) Then
                        Console.WriteLine("Next results")
                        driver.FindElement(By.CssSelector("tr.copysmall > td:nth-child(2) > a:nth-child(1)")).Click()
                        Threading.Thread.Sleep(rnd.Next(123, 1234))
                    Else
                        Console.WriteLine("No more results") : Exit While
                    End If
                End While
                Dim ds As New DataSet

                Dim dtChatSessions As DataTable = Helper.ListToDataTable(ChatSessions)
                dtChatSessions.TableName = "Chat Sessions"
                For Each r As DataRow In dtChatSessions.Rows
                    Dim TS() As String = Split(r.Item("Start").ToString(), vbCrLf)
                    r.Item("Start") = Date.ParseExact(TS(1), "MM/dd/yyyy", New Globalization.CultureInfo("en-US")).ToString("yyyy/MM/dd") + " " + TS(0)
                Next
                Dim dtChatGenInfoList As DataTable = Helper.ListToDataTable(ChatGenInfoList)
                dtChatGenInfoList.TableName = "General Info"
                Dim dtTran As DataTable = Helper.ListToDataTable(ChatTranscriptList)
                dtTran.TableName = "Transcripts"
                Dim dtNavi As DataTable = Helper.ListToDataTable(NavigationList)
                dtNavi.TableName = "Navigation"
                ds.Tables.Add(dtChatSessions) : ds.Tables.Add(dtChatGenInfoList) : ds.Tables.Add(dtTran) : ds.Tables.Add(dtNavi)

                'NPOIXlsUtil.RenderDataSetToExcel(ds, _
                'String.Format("LivePerson_{0}_Sessions_{1}-{2}_SavedOn_{3}.xls", SelectedSessionType.ToString(), FromDate.ToString("yyyyMMddHHmm"), ToDate.ToString("yyyyMMddHHmm"), Now.ToString("yyyyMMddHHmm")))


                With LivePersonDataSet1
                    .Sessions = ChatSessions : .GenInfoRec = ChatGenInfoList : .Transcripts = ChatTranscriptList : .Navigations = NavigationList
                End With

            Else
                Console.WriteLine("Cannot get search result!")
                recall += 1
                driver.Close()
                StartProcedure()
                'Console.Read()
            End If
            Return LivePersonDataSet1
        Catch ex As Exception
            recall += 1
            driver.Close()
            Threading.Thread.Sleep(60 * 1000)
            StartProcedure()
        End Try
        
    End Function

    Function LoginLivePerson() As FirefoxDriver
        'ID: #s_swepi_1, PWD: #s_swepi_2
        Console.WriteLine("Logging in LivePerson")
        Dim driver As New FirefoxDriver()
        Try
            driver.Navigate.GoToUrl("https://server.iad.liveperson.net/hc/s-68676965/web/m-LA/members/wsuk-5902627579886552170/home.jsp?ac=true")
            driver.FindElement(By.CssSelector("#siteNumber")).SendKeys(LiveChatID)
            driver.FindElement(By.CssSelector("#userName")).SendKeys(ID)
            driver.FindElement(By.CssSelector("#sitePass")).SendKeys(PWD)
            Threading.Thread.Sleep(2266)
            '.login_button
            'If WebDriverUtil.IsElementReady(driver, By.Id("login-button primary"), 10) Then
            '    driver.FindElement(By.Id("login-button primary")).Click()
            'Else
            '    MsgBox("cannot locate login button")
            'End If

            'Threading.Thread.Sleep(6000)
            'If WebDriverUtil.IsElementReady(driver, By.Id("login-button primary"), 3) Then
            '    driver.FindElement(By.Id("login-button primary")).Click()
            'End If
            If WebDriverUtil.IsElementReady(driver, By.CssSelector(".login-button.primary"), 10) Then
                driver.FindElement(By.CssSelector(".login-button.primary")).Click()
            Else
                MsgBox("cannot locate login button")
            End If

            Threading.Thread.Sleep(6000)
            If WebDriverUtil.IsElementReady(driver, By.CssSelector(".login-button.primary"), 3) Then
                driver.FindElement(By.CssSelector(".login-button.primary")).Click()
            End If

            'driver.FindElement(By.CssSelector("#login-button primary")).Click()
            'driver.FindElement(By.ClassName("login_button")).Click()

            Console.WriteLine("Logged in")
        Catch ex As Exception
            recall += 1
            driver.Close()
            Threading.Thread.Sleep(60 * 1000)
            StartProcedure()
        End Try
        
        Return driver
    End Function

    Function GetChatDetail(ByRef driver As FirefoxDriver) As ChatDetails
        Dim ChatDetails1 As New ChatDetails()
        If WebDriverUtil.IsElementReady(driver, By.CssSelector(".bkgdeditTable > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(2) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(1)"), 3) Then
            Dim generalInfoTable = driver.FindElement(By.CssSelector(".bkgdeditTable > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(2) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(1)"))
            Dim dataRows = generalInfoTable.FindElements(By.CssSelector("tr"))
            Console.WriteLine("Getting Chat Detail Info...")
            For Each trRow In dataRows
                Dim TDs = trRow.FindElements(By.TagName("td"))
                If TDs.Count = 2 Then
                    Dim DataValue As String = TDs(1).Text
                    Select Case TDs(0).Text
                        Case "Chat start time"
                            ChatDetails1.GenInfo.StartTime = DataValue
                        Case "Chat end time"
                            ChatDetails1.GenInfo.EndTime = DataValue
                        Case "Duration (actual chatting time)"
                            ChatDetails1.GenInfo.Duration = DataValue
                        Case "Chat starting page"
                            ChatDetails1.GenInfo.StartingPage = DataValue
                        Case "Operator"
                            ChatDetails1.GenInfo.Operator = DataValue
                        Case "Browser"
                            ChatDetails1.GenInfo.Browser = DataValue
                        Case "Operating System"
                            ChatDetails1.GenInfo.OS = DataValue
                        Case "Host address"
                            ChatDetails1.GenInfo.HostAddr = DataValue
                        Case "Host IP"
                            ChatDetails1.GenInfo.IP = DataValue
                        Case "Session Referrer"
                            ChatDetails1.GenInfo.Referrer = DataValue
                        Case "Real Time Session ID"
                            ChatDetails1.GenInfo.RealTimeSessionID = DataValue : ChatDetails1.RealTimeSessionID = DataValue
                            Console.WriteLine("{0}:{1}", TDs(0).Text, DataValue)
                        Case "Country"
                            ChatDetails1.GenInfo.Country = DataValue
                        Case "City"
                            ChatDetails1.GenInfo.City = DataValue
                        Case "Organization"
                            ChatDetails1.GenInfo.Org = "" 'DataValue
                        Case "World Region"
                            ChatDetails1.GenInfo.WorldRegion = DataValue
                        Case "Time Zone"
                            ChatDetails1.GenInfo.Timezone = DataValue
                        Case "ISP"
                            ChatDetails1.GenInfo.ISP = DataValue
                        Case "Postal Code"
                            ChatDetails1.GenInfo.PostalCode = DataValue
                        Case "Connection Type"
                            ChatDetails1.GenInfo.ConnectionType = DataValue
                        Case "Chat Indicator"
                            ChatDetails1.GenInfo.ChatIndicator = DataValue
                        Case Else
                            Console.WriteLine("Not handle {0}:{1}", TDs(0).Text, DataValue)
                    End Select

                End If
            Next
        End If

        If WebDriverUtil.IsElementReady(driver, By.CssSelector("body > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(2) > table > tbody > tr:nth-child(6) > td > table > tbody > tr:nth-child(1) > td:nth-child(2) > table.bkgdeditTable > tbody > tr > td > table:nth-child(5) > tbody > tr > td > table > tbody > tr:nth-child(3) > td"), 3) Then
            Dim TranTable = driver.FindElement(By.CssSelector("body > table:nth-child(2) > tbody > tr:nth-child(1) > td:nth-child(2) > table > tbody > tr:nth-child(6) > td > table > tbody > tr:nth-child(1) > td:nth-child(2) > table.bkgdeditTable > tbody > tr > td > table:nth-child(5) > tbody > tr > td > table > tbody > tr:nth-child(3) > td"))

            driver.FindElement(By.CssSelector(".bkgdeditTable > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(5) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > form:nth-child(2) > input:nth-child(1)")).Click()
            Dim convRows = TranTable.FindElements(By.ClassName("tabledata"))
            Console.WriteLine("Transcript of sessionid:{0}", ChatDetails1.GenInfo.RealTimeSessionID)
            For Each cRow In convRows
                Dim convRowTDs = cRow.FindElements(By.TagName("td"))
                If convRowTDs.Count = 2 AndAlso convRowTDs(0).FindElement(By.TagName("div")).GetAttribute("id").EndsWith("timestamp") Then
                    Dim cs As New ChatTranscript
                    cs.RealTimeSessionID = ChatDetails1.RealTimeSessionID
                    cs.Timestamp = convRowTDs(0).FindElement(By.TagName("div")).Text.Substring(1)
                    cs.Timestamp = cs.Timestamp.Substring(0, cs.Timestamp.Length - 1)
                    cs.Name = convRowTDs(1).FindElement(By.TagName("strong")).Text : cs.Name = cs.Name.Substring(0, cs.Name.Length - 3)
                    cs.Text = convRowTDs(1).Text.Substring(convRowTDs(1).Text.IndexOf(":") + 2)
                    ChatDetails1.ChatTranscripts.Add(cs)
                    'Console.WriteLine("Dialog: {0}-{1}", cs.Name, cs.Text)
                End If
            Next
        End If

        'If WebDriverUtil.IsElementReady(driver, By.CssSelector(".bkgdeditTable > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(8)"), 3) Then
        '    Dim NavigationTable = driver.FindElement(By.CssSelector(".bkgdeditTable > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(8)"))
        '    Dim NaviRows = NavigationTable.FindElements(By.TagName("tr"))
        '    For Each NaviTr In NaviRows
        '        If NaviTr.FindElements(By.TagName("td")).Count = 3 AndAlso NaviTr.GetAttribute("class") = "tabledata" Then
        '            Dim TDs = NaviTr.FindElements(By.TagName("td"))
        '            If TDs(1).Text.ToLower().StartsWith("http") Then
        '                Dim navigation1 As New Navigation
        '                navigation1.SeqNo = TDs(0).Text : navigation1.URL = TDs(1).Text : navigation1.Timestamp = TDs(2).Text : navigation1.RealTimeSessionID = ChatDetails1.RealTimeSessionID
        '                ChatDetails1.Navigations.Add(navigation1)
        '            End If

        '        End If
        '    Next
        'End If
        Dim NaviCandidateTDs = driver.FindElements(By.TagName("td"))
        For Each NaviCandTD In NaviCandidateTDs
            Try
                If NaviCandTD.Text = "Navigation" Then

                End If
            Catch ex As WebDriverException
                Console.WriteLine("Time out exception occurred when getting 'NaviCandTD.Text'")
                Exit For
            End Try
            If NaviCandTD.Text = "Navigation" Then
                'WebDriverUtil.GetParent(WebDriverUtil.GetParent(WebDriverUtil.GetParent(NaviCandTD)))
                Dim NaviRows = WebDriverUtil.GetParent(WebDriverUtil.GetParent(WebDriverUtil.GetParent(NaviCandTD))).FindElements(By.TagName("tr"))
                For Each NaviTr In NaviRows
                    If NaviTr.FindElements(By.TagName("td")).Count = 3 AndAlso NaviTr.GetAttribute("class") = "tabledata" Then
                        Dim TDs = NaviTr.FindElements(By.TagName("td"))
                        If TDs(1).Text.ToLower().StartsWith("http") Then
                            Dim navigation1 As New Navigation
                            navigation1.SeqNo = TDs(0).Text
                            Try
                                navigation1.URL = TDs(1).FindElement(By.TagName("a")).GetAttribute("href")
                            Catch ex As Exception
                                navigation1.URL = ""
                            End Try

                            navigation1.Timestamp = TDs(2).Text
                            navigation1.RealTimeSessionID = ChatDetails1.RealTimeSessionID
                            ChatDetails1.Navigations.Add(navigation1)
                        End If
                    End If
                Next

            End If
            If NaviCandTD.Text = "Pre-Chat Survey" Then
                Dim NaviRows = WebDriverUtil.GetParent(WebDriverUtil.GetParent(WebDriverUtil.GetParent(NaviCandTD))).FindElements(By.TagName("tr"))
                Dim CompanyRow = Nothing, CountryRow = Nothing
                If NaviRows.Count = 6 Then
                    CompanyRow = NaviRows(4)
                    Try
                        CountryRow = NaviRows(5)
                    Catch ex As Exception

                    End Try
                End If
                If NaviRows.Count = 5 Then
                    CompanyRow = NaviRows(3)
                    Try
                        CountryRow = NaviRows(4)
                    Catch ex As Exception

                    End Try
                End If
                If CompanyRow IsNot Nothing AndAlso CompanyRow.FindElements(By.TagName("td")).Count = 3 AndAlso CompanyRow.GetAttribute("class") = "tabledata" Then
                    Dim TDs = CompanyRow.FindElements(By.TagName("td"))
                    ChatDetails1.GenInfo.Org = TDs(1).Text
                End If
                If CountryRow IsNot Nothing AndAlso CountryRow.FindElements(By.TagName("td")).Count = 3 AndAlso CountryRow.GetAttribute("class") = "tabledata" Then
                    Dim TDs = CountryRow.FindElements(By.TagName("td"))
                    ChatDetails1.GenInfo.Country1 = TDs(1).Text.ToString.Trim
                End If

                Exit For
            End If
        Next

        Return ChatDetails1
    End Function

    Public Class ChatDetails
        Public Property RealTimeSessionID As String
        Public Property GenInfo As GeneralChatInfo : Public Property ChatTranscripts As List(Of ChatTranscript) : Public Property PreCS As PreChatSurvey : Public Property CV As CustomVariables
        Public Property Navigations As List(Of Navigation)
        Public Sub New()
            RealTimeSessionID = "" : GenInfo = New GeneralChatInfo() : ChatTranscripts = New List(Of ChatTranscript) : PreCS = New PreChatSurvey() : CV = New CustomVariables() : Navigations = New List(Of Navigation)
        End Sub
    End Class

    Public Enum SessionType
        Chats
        Calls
    End Enum

    Public Class LivePersonDataSet
        Public Property Sessions As List(Of Session) : Public Property GenInfoRec As List(Of GeneralChatInfo)
        Public Property Transcripts As List(Of ChatTranscript) : Public Property Navigations As List(Of Navigation)
        Public Sub New()
            Sessions = New List(Of Session) : GenInfoRec = New List(Of GeneralChatInfo) : Transcripts = New List(Of ChatTranscript) : Navigations = New List(Of Navigation)
        End Sub
    End Class

    Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal Subject As String, ByVal Body As String, ByVal IsBodyHtml As Boolean, _
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
