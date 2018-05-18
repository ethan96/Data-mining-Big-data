Imports SyncInvalidEmail.AEUOWA
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml.Serialization
Imports System.Xml
Imports System.Web.Services.Protocols
Imports System.IO
Imports WatiN.Core
Imports System.Text

Module Module1

    Sub Main()
        Console.WriteLine("SyncAlibabaEmail is Running...")
        Dim CuApi As New CurationPoolApi.IntApi
        CuApi.Url = "http://unica.advantech.com.tw/Services/IntApi.asmx"

        'Dim CuApi As New CPool_Frank.IntApi


        While True
            Try
                Dim esb As New MyExchangeServiceBinding()
                'esb.Credentials = New NetworkCredential("ebiz.aeu", "@myadvantech78!", "aesc_nt")
                'esb.Url = "https://mail.advantech.eu/EWS/Exchange.asmx"
                esb.Credentials = New NetworkCredential("ebiz.aeu@advantech.com", "@myadvantech78!")
                'esb.UseDefaultCredentials = True
                esb.Url = "https://outlook.office365.com/EWS/Exchange.asmx"



                'esb.RequestServerVersionValue = New RequestServerVersion()
                'esb.RequestServerVersionValue.Version = ExchangeVersionType.Exchange2007_SP1
                esb.Timeout = -1
                ServicePointManager.ServerCertificateValidationCallback = Function(sender As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As Net.Security.SslPolicyErrors) True

                Dim findItemRequest As New FindItemType()
                findItemRequest.Traversal = ItemQueryTraversalType.Shallow

                Dim itemProperties As New ItemResponseShapeType()
                itemProperties.BaseShape = DefaultShapeNamesType.AllProperties

                findItemRequest.ItemShape = itemProperties

                Dim folderIDArray As DistinguishedFolderIdType() = New DistinguishedFolderIdType(0) {}
                folderIDArray(0) = New DistinguishedFolderIdType() : folderIDArray(0).Id = DistinguishedFolderIdNameType.inbox

                Dim fieldorder As New FieldOrderType(), findindex As New PathToUnindexedFieldType()
                findindex.FieldURI = UnindexedFieldURIType.itemDateTimeReceived : fieldorder.Item = findindex : fieldorder.Order = SortDirectionType.Ascending

                findItemRequest.SortOrder = New FieldOrderType(0) {}
                findItemRequest.SortOrder(0) = fieldorder : findindex = New PathToUnindexedFieldType() : findindex.FieldURI = UnindexedFieldURIType.messageIsRead

                Dim fieldconst As New FieldURIOrConstantType(), constvalue As New ConstantValueType()
                constvalue.Value = "0" : fieldconst.Item = constvalue

                Dim isread As New IsEqualToType()
                isread.Item = findindex : isread.FieldURIOrConstant = fieldconst

                Dim filter As New RestrictionType()
                filter.Item = isread

                'findItemRequest.Restriction = filter
                findItemRequest.ParentFolderIds = folderIDArray


                Dim sb As New StringBuilder
                sb.AppendLine("SyncAlibabaEmail is Running..." + Now.ToString + "<br/>")
                Dim IsFetchDirectIndustry As Boolean = False

                Dim fis As FindItemResponseType = esb.FindItem(findItemRequest), rmta As ResponseMessageType() = fis.ResponseMessages.Items
                For Each rmt As ResponseMessageType In rmta
                    'Response.Write(rmt.MessageText + rmt.ResponseCode.ToString() + "<br/>") 
                    Dim firmt As FindItemResponseMessageType = DirectCast(rmt, FindItemResponseMessageType)
                    If firmt.ResponseClass = ResponseClassType.Success Then
                        Dim fit As FindItemParentType = firmt.RootFolder, obj As Object = fit.Item

                        If TypeOf obj Is ArrayOfRealItemsType Then
                            Dim msgitems As ArrayOfRealItemsType = DirectCast(obj, ArrayOfRealItemsType)
                            If obj Is Nothing OrElse msgitems Is Nothing OrElse msgitems.Items Is Nothing Then
                                Exit For
                            End If
                            For iMsg As Integer = 0 To msgitems.Items.Length - 1

                                If TypeOf msgitems.Items(iMsg) Is MessageType Then

                                    'check if the message has been not read 
                                    Dim omsg As MessageType = TryCast(msgitems.Items(iMsg), MessageType), emailid As String = omsg.ItemId.Id
                                    Dim _IsProcessedEmail As Boolean = CuApi.IsProcessedAlibabaEmail(emailid)
                                    'Frank:Processing email that be sent after 20121217
                                    Dim _EmailCreatedDate As Integer = Integer.Parse(Format(omsg.DateTimeCreated, "yyyyMMdd"))
                                    If _EmailCreatedDate < 20121217 Then
                                        _IsProcessedEmail = True
                                        'Console.WriteLine("Skip this email because the created date befor 2012-12-17...")
                                    End If

                                    'If Not omsg.IsRead Or True Then
                                    If Not _IsProcessedEmail Then

                                        Dim getrequest As GetItemType = getitemrequest(omsg.ItemId.Id)
                                        Dim getresponse As New GetItemResponseType()
                                        Dim getOk As Boolean = False
                                        Try
                                            getresponse = esb.GetItem(getrequest)
                                            getOk = True
                                        Catch e3 As Exception
                                            getOk = False
                                        End Try
                                        If getOk Then
                                            Dim getrmta As ResponseMessageType() = getresponse.ResponseMessages.Items
                                            For Each getrmt As ResponseMessageType In getrmta
                                                Dim iteminfo As ItemInfoResponseMessageType = DirectCast(getrmt, ItemInfoResponseMessageType)

                                                If iteminfo.ResponseClass = ResponseClassType.Success Then
                                                    Dim arrmsgs As ArrayOfRealItemsType = DirectCast(iteminfo.Items, ArrayOfRealItemsType)
                                                    For Each msg As MessageType In arrmsgs.Items
                                                        'Dim ws As New EC.EC()
                                                        'ws.UseDefaultCredentials = True : ws.Timeout = -1
                                                        If (msg.Sender IsNot Nothing AndAlso msg.Sender.Item IsNot Nothing) Or (msg.From IsNot Nothing AndAlso msg.From.Item IsNot Nothing) Then
                                                            Dim sEmailFrom As EmailAddressType = If(msg.Sender Is Nothing, msg.From.Item, msg.Sender.Item)
                                                            'If (msg.Sender Is Nothing) Then
                                                            '    sEmailFrom = msg.From.Item
                                                            'Else
                                                            '    sEmailFrom = msg.Sender.Item
                                                            'End If
                                                            If sEmailFrom.EmailAddress IsNot Nothing Then
                                                                'Dim sFrom As String = sEmailFrom.EmailAddress.ToLower()
                                                                'MsgBox(msg.Body.Value)
                                                                Console.WriteLine("==============")
                                                                Console.WriteLine("Email=" & sEmailFrom.EmailAddress)
                                                                Dim blProcEmailFlag As Boolean = False, _errmsg As String = String.Empty

                                                                '***Identify email source type
                                                                Select Case GetServiceType(msg)

                                                                    Case ServiceType.Alibaba
                                                                        Try
                                                                            Console.WriteLine("call ProcessAlibabaEmail for " & sEmailFrom.EmailAddress)
                                                                            blProcEmailFlag = CuApi.ProcessAlibabaEmail(msg.Body.Value, emailid)
                                                                        Catch ex As Exception
                                                                            Console.WriteLine("call ProcessAlibabaEmail error...")
                                                                        End Try

                                                                    Case ServiceType.Traceparts

                                                                        If _EmailCreatedDate < 20130318 Then
                                                                            _IsProcessedEmail = True : Continue For
                                                                        End If

                                                                        Try
                                                                            Console.WriteLine("call ProcessTracepartsExcel for " & sEmailFrom.EmailAddress)
                                                                            Dim _AttachBytes() As Byte = GetAttachmentInByte(msg, esb)
                                                                            If _AttachBytes IsNot Nothing Then
                                                                                'Process excel
                                                                                blProcEmailFlag = CuApi.ProcessTracepartsExcel(_AttachBytes, emailid, _errmsg)
                                                                            End If
                                                                            _AttachBytes = Nothing
                                                                        Catch ex As Exception
                                                                            Console.WriteLine("call ProcessTracepartsExcel error...")
                                                                        End Try

                                                                    Case ServiceType.Traceparts_SalesInquiry

                                                                        If _EmailCreatedDate < 20130418 Then
                                                                            _IsProcessedEmail = True : Continue For
                                                                        End If

                                                                        Try
                                                                            Console.WriteLine("call ProcessTracepartsExcel for " & sEmailFrom.EmailAddress)
                                                                            'Process email
                                                                            blProcEmailFlag = CuApi.ProcessTracepartEmail(msg.Body.Value, emailid, _errmsg)
                                                                        Catch ex As Exception
                                                                            Console.WriteLine("call ProcessTracepartsExcel error...")
                                                                        End Try


                                                                    Case ServiceType.Globalspec

                                                                        If _EmailCreatedDate < 20130318 Then
                                                                            _IsProcessedEmail = True : Continue For
                                                                        End If

                                                                        Try
                                                                            Console.WriteLine("call ProcessGlobalspecExcel for " & sEmailFrom.EmailAddress)
                                                                            Dim _AttachBytes() As Byte = GetAttachmentInByte(msg, esb)
                                                                            'Process excel
                                                                            blProcEmailFlag = CuApi.ProcessGlobalspecExcel(_AttachBytes, emailid, _errmsg)
                                                                            _AttachBytes = Nothing
                                                                        Catch ex As Exception
                                                                            Console.WriteLine("call ProcessGlobalspecExcel error...")
                                                                        End Try

                                                                    Case ServiceType.Globalspec_SalesInquiry
                                                                        If _EmailCreatedDate < 20130409 Then
                                                                            _IsProcessedEmail = True
                                                                            Console.WriteLine("Email has been ignored because receive date is earlier than 2013-04-09.")
                                                                            Continue For
                                                                        End If

                                                                        Dim _body As String = msg.Body.Value
                                                                        Try
                                                                            Console.WriteLine("call ProcessGlobalspecEmail for " & sEmailFrom.EmailAddress)
                                                                            blProcEmailFlag = CuApi.ProcessGlobalspecEmail(msg.Body.Value, emailid, _errmsg)
                                                                        Catch ex As Exception
                                                                            Console.WriteLine("call ProcessAlibabaEmail error...")
                                                                        End Try

                                                                    Case ServiceType.Directindustry
                                                                        If _EmailCreatedDate < 20130412 Then
                                                                            _IsProcessedEmail = True : Continue For
                                                                        End If
                                                                        Dim _body As String = msg.Body.Value
                                                                        Try
                                                                            IsFetchDirectIndustry = True
                                                                            sb.AppendLine("Start calling ProcessDirectIndustrySalesInquiryEmail sales inquiry email for " + sEmailFrom.EmailAddress + "..." + Now.ToString + "<br/>")
                                                                            Console.WriteLine("call ProcessDirectIndustrySalesInquiryEmail sales inquiry email for " & sEmailFrom.EmailAddress)
                                                                            blProcEmailFlag = CuApi.ProcessDirectIndustrySalesInquiryEmail(msg.Body.Value, emailid, _errmsg)
                                                                            sb.AppendLine("End calling ProcessDirectIndustrySalesInquiryEmail sales inquiry email for " + sEmailFrom.EmailAddress + "..." + Now.ToString + "<br/>")
                                                                        Catch ex As Exception
                                                                            Console.WriteLine("calling ProcessDirectIndustrySalesInquiryEmail sales inquiry email error...")
                                                                        End Try

                                                                    Case ServiceType.MedicalExpo
                                                                        If _EmailCreatedDate < 20140815 Then
                                                                            _IsProcessedEmail = True : Continue For
                                                                        End If
                                                                        Dim _body As String = msg.Body.Value
                                                                        Try
                                                                            Console.WriteLine("call ProcessMedicalExpoSalesInquiryEmail sales inquiry email for " & sEmailFrom.EmailAddress)
                                                                            blProcEmailFlag = CuApi.ProcessMedicalExpoSalesInquiryEmail(msg.Body.Value, emailid, _errmsg)
                                                                        Catch ex As Exception
                                                                            Console.WriteLine("call ProcessMedicalExpoSalesInquiryEmail sales inquiry email error...")
                                                                        End Try
                                                                    Case ServiceType.SolarACN
                                                                        CreateActivity(msg)
                                                                    Case Else
                                                                        'Console.WriteLine("This mail was not sent by Alibaba, Traceparts, MedicalExpo, Globalspec or Directindustry.")

                                                                End Select

                                                                Console.WriteLine("Processed statue=" & blProcEmailFlag.ToString())
                                                                If Not String.IsNullOrEmpty(_errmsg) Then Console.WriteLine("Processed message=" & _errmsg)

                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    End If

                                    'Frank:2013/03/28:Remove old(1 month ago) email
                                    Console.WriteLine("Email Item Received date=" & omsg.DateTimeReceived)

                                    If DateDiff(DateInterval.Day, omsg.DateTimeReceived, Now.Date) > 20 Then
                                        'Console.WriteLine("Email will be deleted(" & DateDiff(DateInterval.Month, msg.DateTimeReceived, Now.Date) & ")")
                                        Console.WriteLine(omsg.Subject)
                                        DeleteItem(esb, emailid)
                                        Console.WriteLine("Email has been deleted")
                                        'Exit For
                                    End If

                                End If
                            Next
                        End If
                    End If
                Next

                sb.AppendLine("SyncAlibabaEmail is completed..." + Now.ToString + "<br/>")
                'If IsFetchDirectIndustry Then SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Sync DirectIndustry Email Message", sb.ToString, True)
            Catch ex As Exception
                'Dim BCopy As New SqlClient.SqlBulkCopy(strRFM)
                'BCopy.DestinationTableName = "INVALID_EMAIL"
                'BCopy.WriteToServer(dt)
                Console.WriteLine(ex.ToString)
                'SendEmail("tc.chen@advantech.eu,rudy.wang@advantech.com.tw", "tc.chen@advantech.eu", "Sync Invalid Email error", ex.ToString, True)
                CenterLibrary.MailUtil.SendEmail("frank.chung@advantech.com.tw,rudy.wang@advantech.com.tw", "tc.chen@advantech.eu", "Sync Invalid Email error", ex.ToString, True, "", "")
            End Try
            'Console.Read()
            'Frank 2013/04/25
            'Startup this program by OS schedule job
            Exit Sub
            Dim _NextRunMin As Integer = 5
            Console.WriteLine("Sleeping...")
            Console.WriteLine("Next running time will be " & DateAdd(DateInterval.Minute, _NextRunMin, Now))
            Threading.Thread.Sleep(_NextRunMin * 60 * 1000)
        End While

    End Sub

    Private Function GetAttachmentInByte(ByVal msg As MessageType, ByVal esb As ExchangeServiceBinding) As Byte()

        '取得目前郵件的附檔集合
        If msg.HasAttachments Then

            Dim attachments() As AttachmentType = msg.Attachments

            '循環取得的每一郵件的附件
            For i As Integer = 0 To attachments.Length - 1

                '定義GetAttachmentType，設定相對應的屬性，進行模式驗證來取得附件
                Dim getAttachment As GetAttachmentType = New GetAttachmentType()
                Dim attachmentIDArry() As RequestAttachmentIdType = New RequestAttachmentIdType() {Nothing}
                attachmentIDArry(0) = New RequestAttachmentIdType()
                attachmentIDArry(0).Id = attachments(i).AttachmentId.Id
                getAttachment.AttachmentIds = attachmentIDArry
                getAttachment.AttachmentShape = New AttachmentResponseShapeType()

                '取得附件
                Dim getAttachmentResponse As GetAttachmentResponseType = esb.GetAttachment(getAttachment)
                Dim ccc = CType(CType(getAttachmentResponse.ResponseMessages.Items(0), AttachmentInfoResponseMessageType).Attachments(0), AttachmentType)
                If ccc.GetType.ToString.Equals("SyncInvalidEmail.AEUOWA.FileAttachmentType") Then

                    Return DirectCast(ccc, SyncInvalidEmail.AEUOWA.FileAttachmentType).Content

                    'Dim filearry() As Byte = DirectCast(ccc, SyncInvalidEmail.AEUOWA.FileAttachmentType).Content
                    ''If attachment is zip file
                    'If ccc.Name.EndsWith(".zip", StringComparison.InvariantCultureIgnoreCase) Then

                    '    Dim s As MemoryStream = New MemoryStream(filearry)
                    '    Dim filename As String = String.Empty
                    '    'Unzip
                    '    Dim s1 As MemoryStream = ExtractFileFromZIP(s, filename)

                    '    Dim fs As New FileStream("D:\EmailAttachment\" & filename, FileMode.Create)
                    '    fs.Write(s1.ToArray, 0, s1.Length)
                    '    fs.Flush()
                    '    fs.Close()
                    'Else
                    '    Dim fs As New FileStream("D:\EmailAttachment\" & ccc.Name, FileMode.Create)
                    '    fs.Write(filearry, 0, filearry.Length)
                    '    fs.Flush()
                    '    fs.Close()
                    'End If
                Else

                End If
            Next

        End If
        Return Nothing
    End Function

    Private Function ExtractFileFromZIP(ByRef ms As IO.Stream, ByRef FileName As String) As IO.Stream
        Dim zin As New ICSharpCode.SharpZipLib.Zip.ZipInputStream(ms)
        Dim zEntry As ICSharpCode.SharpZipLib.Zip.ZipEntry = Nothing
        While True
            Try
                zEntry = zin.GetNextEntry()
            Catch ex As ICSharpCode.SharpZipLib.Zip.ZipException
                Exit While
            End Try
            If zEntry IsNot Nothing Then
                If zEntry.IsFile Then
                    FileName = zEntry.Name
                    Dim ftype As ValidFileTypes = IsValidFileType(zEntry.Name)
                    If ftype <> ValidFileTypes.NOT_SUPPORTED Then
                        Dim bs As Byte() = New Byte(ms.Length) {}
                        Dim extMS As New IO.MemoryStream
                        ICSharpCode.SharpZipLib.Core.StreamUtils.Copy(zin, extMS, bs)
                        If bs IsNot Nothing AndAlso bs.Length > 0 Then
                            extMS.Position = 0
                            Return extMS
                        End If
                    End If
                End If
            End If
        End While
        Return Nothing
    End Function

    Public Enum ValidFileTypes
        PDF
        DOC
        DOCX
        XLS
        XLSX
        PPT
        PPTX
        RAR
        ZIP
        TXT
        NOT_SUPPORTED
    End Enum


    Private Function IsValidFileType(ByVal FileName As String) As ValidFileTypes
        If FileName.Contains(".") Then
            Dim ext As String = FileName.Substring(FileName.LastIndexOf(".") + 1).ToUpper()
            Select Case ext
                Case "PDF"
                    Return ValidFileTypes.PDF
                Case "DOC"
                    Return ValidFileTypes.DOC
                Case "DOCX"
                    Return ValidFileTypes.DOCX
                Case "PPT"
                    Return ValidFileTypes.PPT
                Case "PPTX"
                    Return ValidFileTypes.PPTX
                Case "XLS"
                    Return ValidFileTypes.XLS
                Case "XLSX"
                    Return ValidFileTypes.XLSX
                Case "RAR"
                    Return ValidFileTypes.RAR
                Case "ZIP"
                    Return ValidFileTypes.ZIP
                Case "TXT"
                    Return ValidFileTypes.TXT
                Case Else
                    Return ValidFileTypes.NOT_SUPPORTED
            End Select
        Else
            Return ValidFileTypes.NOT_SUPPORTED
        End If
    End Function

    Private Sub DeleteItem(ByVal esb As ExchangeServiceBinding, ByVal sItemId As String)
        ' Create the request. 
        Dim request As New DeleteItemType()
        ' Identify the items to delete. 
        Dim items As ItemIdType() = New ItemIdType(0) {}
        items(0) = New ItemIdType() : items(0).Id = sItemId : request.ItemIds = items

        ' Identify how deleted items are handled. 
        request.DeleteType = DisposalType.HardDelete
        Try
            ' Send the response and receive the request. 
            Dim response As DeleteItemResponseType = esb.DeleteItem(request), aormt As ArrayOfResponseMessagesType = response.ResponseMessages, rmta As ResponseMessageType() = aormt.Items
        Catch e As Exception
            'Console.WriteLine(e.Message); 
        End Try
    End Sub

    Public Function updateitemrequest(ByVal sItemId As String, ByVal sChangeKey As String) As UpdateItemType
        Dim updateitemrequest2 As New UpdateItemType()

        Dim itemtoupdate As New ItemIdType()
        itemtoupdate.Id = sItemId
        itemtoupdate.ChangeKey = sChangeKey

        'message to be updated 
        Dim msgitem As New MessageType()
        msgitem.IsRead = True
        msgitem.IsReadSpecified = True

        ' Create the set update 
        Dim setItemField As New SetItemFieldType()
        Dim indexfield As New PathToUnindexedFieldType()
        indexfield.FieldURI = UnindexedFieldURIType.messageIsRead

        setItemField.Item = indexfield
        setItemField.Item1 = msgitem

        updateitemrequest2.ItemChanges = New ItemChangeType(0) {}
        Dim itemchange As New ItemChangeType()
        itemchange.Item = itemtoupdate
        itemchange.Updates = New ItemChangeDescriptionType(0) {}
        itemchange.Updates(0) = setItemField
        updateitemrequest2.ItemChanges(0) = itemchange
        updateitemrequest2.MessageDisposition = MessageDispositionType.SaveOnly
        updateitemrequest2.MessageDispositionSpecified = True
        updateitemrequest2.ConflictResolution = ConflictResolutionType.AlwaysOverwrite

        Return updateitemrequest2
    End Function

    Public Function getitemrequest(ByVal sItemId As String) As GetItemType

        Dim getrequest As New GetItemType()
        Dim itemproperties As New ItemResponseShapeType()
        'itemproperties.BaseShape = DefaultShapeNamesType.AllProperties : itemproperties.BodyType = BodyTypeResponseType.HTML
        itemproperties.BaseShape = DefaultShapeNamesType.AllProperties
        itemproperties.BodyType = BodyTypeResponseType.Best
        itemproperties.BodyTypeSpecified = True
        itemproperties.IncludeMimeContent = True
        getrequest.ItemShape = itemproperties

        Dim itemtosend As ItemIdType() = New ItemIdType(0) {}
        itemtosend(0) = New ItemIdType()
        itemtosend(0).Id = sItemId

        getrequest.ItemIds = itemtosend
        Return getrequest
    End Function

    Enum ServiceType
        Alibaba = 0
        Traceparts = 1
        Globalspec = 2
        Directindustry = 3
        Globalspec_SalesInquiry = 4
        Traceparts_SalesInquiry = 5
        MedicalExpo = 6
        SolarACN = 7
    End Enum

    Public Function GetServiceType(ByVal msg As MessageType) As ServiceType

        Dim sEmailSender As EmailAddressType = If(msg.Sender Is Nothing, msg.From.Item, msg.Sender.Item)
        Dim sEmailSubject As String = msg.ConversationTopic
        Dim sEmailBody As String = msg.Body.Value
        Dim sEmailTo = msg.ToRecipients

        If sEmailSubject.StartsWith("[GLOBALSPEC]", StringComparison.InvariantCultureIgnoreCase) _
            AndAlso sEmailBody.IndexOf(" GLOBALSPEC (HTTP://WWW.GLOBALSPEC.COM). ", StringComparison.InvariantCultureIgnoreCase) > 0 Then
            Return ServiceType.Globalspec_SalesInquiry
        End If

        If sEmailSubject.StartsWith("[IHS ENGINEERING360]", StringComparison.InvariantCultureIgnoreCase) Then
            Return ServiceType.Globalspec_SalesInquiry
        End If

        If sEmailSubject.StartsWith("All Globalspec Contacts", StringComparison.CurrentCultureIgnoreCase) Then
            Return ServiceType.Globalspec
        End If

        If sEmailSubject.IndexOf("Request for quotation received from www.tracepartsonline.net") >= 0 Then
            Return ServiceType.Traceparts_SalesInquiry
        End If


        If sEmailSender.EmailAddress IsNot Nothing Then
            Dim sFrom As String = sEmailSender.EmailAddress.ToLower()
            If sFrom.EndsWith("@service.alibaba.com") Then
                Return ServiceType.Alibaba
            End If
        End If
        If sEmailSender.EmailAddress IsNot Nothing Then
            Dim sFrom As String = sEmailSender.EmailAddress.ToLower()
            If sFrom.EndsWith("@tracepartsonline.net") Then
                Return ServiceType.Traceparts
            End If
        End If
        If sEmailSender.EmailAddress IsNot Nothing Then
            Dim sFrom As String = sEmailSender.EmailAddress.ToLower()
            If sFrom.EndsWith("@globalspec.com") Then
                Return ServiceType.Globalspec
            End If
        End If
        If sEmailSender.EmailAddress IsNot Nothing Then
            Dim sFrom As String = sEmailSender.EmailAddress.ToLower()
            If sFrom.EndsWith("request@directindustry.com") Then
                Return ServiceType.Directindustry
            End If
        End If
        If sEmailSender.EmailAddress IsNot Nothing Then
            Dim sFrom As String = sEmailSender.EmailAddress.ToLower()
            If sFrom.EndsWith("request@medicalexpo.com") Then
                Return ServiceType.MedicalExpo
            End If
        End If

        If sEmailTo IsNot Nothing AndAlso sEmailTo.Count() > 0 Then
            Dim receiver = sEmailTo.Where(Function(m) m.EmailAddress.EndsWith("solar@advantech.com.cn"))
            If (receiver.Count() > 0) Then
                Return ServiceType.SolarACN
            End If
        End If
        Return -1
    End Function


    Public Function IsAlibabaMail(ByVal msg As MessageType) As Boolean

        Dim sEmailFrom As EmailAddressType = msg.Sender.Item
        If sEmailFrom.EmailAddress IsNot Nothing Then
            Dim sFrom As String = sEmailFrom.EmailAddress.ToLower()
            If sFrom.EndsWith("@service.alibaba.com") Then
                Return True
            End If
        End If
        'If msg.Body IsNot Nothing AndAlso msg.Body.Value IsNot Nothing Then
        '    If msg.Body.Value.Contains("www.alibaba.com") Then
        '        'Console.WriteLine(msg.Body.Value)
        '        'Console.Read()
        '        Return True
        '    End If
        'End If
        Return False
    End Function

    Public Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal Subject As String, ByVal Body As String, ByVal IsBodyHtml As Boolean)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml
        mySmtpClient = New Net.Mail.SmtpClient("172.20.0.76")
        mySmtpClient.Send(htmlMessage)
    End Sub

    Public Function GetIEN(ByVal Url As String, ByVal emailid As String, ByRef ErrMsg As String) As Boolean
        Dim CuApi As New CurationPoolApi.IntApi
        CuApi.Url = "http://pis.advantech.com:7100/Services/IntApi.asmx"
        Dim blProcEmailFlag As Boolean = False

        Dim browser As New IE(Url) 'JJ 2014/8/19：開啟IEN主頁面
        browser.WaitForComplete()
        'Dim TableRows As TableRowCollection = browser.Table(Find.ById("EnqGrid")).TableRows()
        Dim TableRows As TableRowCollection = browser.Table(Find.ById("GridHistoric")).TableRows()

        If TableRows.Count > 1 Then 'JJ 2014/8/19：主要下載資料table必須大於1筆資料，1筆主要是title所以必須大於1
            For i As Integer = 1 To TableRows.Count - 1
                Dim detailUrl As String = TableRows(i).TableCells(4).Links(0).Url 'JJ 2014/8/19：取得detail路徑進入Detail頁了
                Dim detail_amount As String = TableRows(i).TableCells(4).Links(0).Text 'JJ 2014/8/19：取得顯示的數字
                If detail_amount <> "0" Then 'JJ 2014/8/19：如果顯示的數字等於0表示沒有資料就不需要進入Detail頁了
                    Dim detailBrowser As New IE(detailUrl) 'JJ 2014/8/19：開啟Detail頁
                    detailBrowser.WaitForComplete()
                    Dim XlsUrl As String = detailBrowser.Link(Find.ById("HLcsv")).Url 'JJ 2014/8/19：取得Excel檔案下載頁路徑
                    Dim xlsBrowser As New IE(XlsUrl) 'JJ 2014/8/19：開啟Excel檔案下載頁
                    xlsBrowser.WaitForComplete()
                    Dim xlsPath As String = xlsBrowser.Link(Find.ById("HLdownload")).Url 'JJ 2014/8/19：取得Excel檔案下載路徑
                    Dim downloadPath As String = "E:\Scheduled_Programs\SyncAlibabaEmail_V2\Files\Excel.xls" 'JJ 2014/8/19：下載到本地端位置
                    My.Computer.Network.DownloadFile(xlsPath, downloadPath, "", "", True, 500, True) 'JJ 2014/8/19：下載

                    Dim dtEXCEL As DataTable = ExcelFile2DataTable(downloadPath) 'JJ 2014/8/19：到本地端位置抓取Excel檔轉為Datatable


                    'blProcEmailFlag = CuApi.ProcessIENSalesInquiryEmail(dtEXCEL, emailid, ErrMsg)
                    xlsBrowser.Close() 'JJ 2014/8/19：關閉Excel下載頁面
                    detailBrowser.Close() 'JJ 2014/8/19：關閉Detail頁
                End If
            Next
        End If

        browser.Close() 'JJ 2014/8/19：關閉最後的IE
        Return True
    End Function

    Public Function ExcelFile2DataTable(ByVal fs As String) As DataTable
        ASPOSECellLicense()
        Dim dt As New DataTable

        Dim wb As New Aspose.Cells.Workbook
        wb.Open(fs)
        For i As Integer = 0 To wb.Worksheets(0).Cells.Count - 1
            If wb.Worksheets(0) IsNot Nothing AndAlso wb.Worksheets(0).Cells(0, i) IsNot Nothing _
             AndAlso wb.Worksheets(0).Cells(0, i).Value IsNot Nothing AndAlso wb.Worksheets(0).Cells(0, i).Value.ToString <> "" Then
                dt.Columns.Add(wb.Worksheets(0).Cells(0, i).Value)
            Else
                Exit For
            End If
        Next
        For i As Integer = 0 To wb.Worksheets(0).Cells.Rows.Count - 1
            Dim r As DataRow = dt.NewRow
            For j As Integer = 0 To dt.Columns.Count - 1
                If Not String.IsNullOrEmpty(wb.Worksheets(0).Cells(i, j).Value) Then
                    r.Item(j) = wb.Worksheets(0).Cells(i, j).Value
                End If

            Next
            dt.Rows.Add(r)
        Next

        Return dt
    End Function

    Public Sub ASPOSECellLicense()
        Try
            Dim license As Aspose.Cells.License = New Aspose.Cells.License()
            Dim strFPath As String = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Aspose.Total.lic")
            license.SetLicense(strFPath)
        Catch ex As Exception
            'Util.SendEmail("jj.lin@advantech.com.tw", "ebiz.aeu@advantech.eu", "MY GLOBAL: setting aspose license error", ex.ToString(), False, "", "")
        End Try
    End Sub

    Public Sub CreateActivity(ByVal msg As MessageType)
        Try

            Dim strErrMsg As String = ""

            Dim dt As New DataTable()
            Dim conn = New SqlClient.SqlConnection(CenterLibrary.DBConnection.MyAdvantechGlobal)
            Dim connMya = New SqlClient.SqlConnection(CenterLibrary.DBConnection.CurationPool)
            Dim cmd = New SqlClient.SqlCommand()
            Dim da = New SqlClient.SqlDataAdapter(String.Format("SELECT COUNT(*) AS COUNT FROM CREATE_ACN_SOLAR_ACTIVITY_LOG WHERE MAILID = '{0}'", msg.ItemId.Id), conn)
            Dim ContactID = ""
            Dim AccountID = ""

            da.Fill(dt)

            If (dt.Rows.Count > 0 AndAlso dt.Rows(0)("COUNT") = "0") Then
                Dim CuApi As New CurationPoolApi.IntApi
                da = New SqlClient.SqlDataAdapter(String.Format("SELECT TOP (1) ROW_ID AS CONTACTID, ACCOUNT_ROW_ID AS ACCOUNTID FROM SIEBEL_CONTACT WHERE EMAIL_ADDRESS = '{0}'", msg.Sender.Item.EmailAddress), connMya)
                dt.Clear()
                da.Fill(dt)

                If (dt.Rows.Count = 1) Then
                    ContactID = dt.Rows(0)("CONTACTID").ToString()
                    AccountID = dt.Rows(0)("ACCOUNTID").ToString()
                End If

                Dim id = CuApi.CreateActivity_New("Email - Inbound", msg.Subject, "Sender Email: " + msg.Sender.Item.EmailAddress + ";" + vbCrLf + "Mail Content: " + msg.Body.Value, AccountID, ContactID, "1-12E3P6Z", "ACN", "MTL", CurationPoolApi.ActivityStatus.In_Progress, "Unica", strErrMsg)
                If (Not String.IsNullOrEmpty(strErrMsg)) Then
                    Throw New Exception("Create Siebel Error " + strErrMsg)
                End If

                conn.Open()
                cmd.Connection = conn
                cmd.CommandText = String.Format("INSERT INTO CREATE_ACN_SOLAR_ACTIVITY_LOG (MAILID,ROWID,CREATEDDATE) VALUES ('{0}','{1}',GETDATE())", msg.ItemId.Id, id)
                Dim result = cmd.ExecuteNonQuery()

            End If
        Catch ex As Exception
            CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw,myadvantech@advantech.com", "rudy.wang@advantech.com.tw", "ACN Solar Create Siebel Activity Error", ex.ToString, True, "", "")
            'SendEmail("myadvantech@advantech.com", "", "Create Siebel Activity Error", ex.ToString, True)
        End Try
    End Sub
End Module
