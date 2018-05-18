Imports SyncInvalidEmail.ACLOWA
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml.Serialization
Imports System.Xml
Imports System.Web.Services.Protocols
Imports System.Web
Imports System.Configuration

Module Module1

    Public strRFM As String = CenterLibrary.DBConnection.MyLocal
    Public strMY As String = CenterLibrary.DBConnection.MyAdvantechGlobal

    Public SoftBounceCode As New List(Of String) From {"4.2.1", "4.5.0", "4.5.1", "4.5.2", "5.2.0", "5.2.1", "5.2.2", "5.3.1", "5.4.5", "5.5.3"}

    Sub Main()
        Dim dt As New DataTable
        With dt.Columns
            .Add("EMAIL") : .Add("INS_DATE") : .Add("SUBJECT") : .Add("EMAIL_BODY") : .Add("REASON_FLAG") : .Add("TYPE") : .Add("CODE") : .Add("MESSAGE")
        End With

        GetSendGridBounceMail(dt, ConfigurationManager.AppSettings("SendGridUID"), ConfigurationManager.AppSettings("SendGridPWD"))
        GetSendGridBounceMail(dt, ConfigurationManager.AppSettings("SendGridUID1"), ConfigurationManager.AppSettings("SendGridPWD1"))
        'GetDataFromWebpower.GetHardBounceMail(dt)

        While True

            Try
                Dim esb As New ExchangeServiceBinding()
                esb.Credentials = New NetworkCredential("edm.advantech", "P@ssw0rd")
                esb.Url = "https://172.20.2.122/EWS/Exchange.asmx"
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
                folderIDArray(0) = New DistinguishedFolderIdType()
                folderIDArray(0).Id = DistinguishedFolderIdNameType.inbox
                'folderIDArray(0).Id = DistinguishedFolderIdNameType.deleteditems
                'folderIDArray(0).Id = DistinguishedFolderIdNameType.junkemail

                Dim fieldorder As New FieldOrderType()
                Dim findindex As New PathToUnindexedFieldType()
                findindex.FieldURI = UnindexedFieldURIType.itemDateTimeReceived
                fieldorder.Item = findindex
                fieldorder.Order = SortDirectionType.Ascending

                findItemRequest.SortOrder = New FieldOrderType(0) {}
                findItemRequest.SortOrder(0) = fieldorder
                findindex = New PathToUnindexedFieldType()
                findindex.FieldURI = UnindexedFieldURIType.messageIsRead

                Dim fieldconst As New FieldURIOrConstantType()
                Dim constvalue As New ConstantValueType()
                constvalue.Value = "0"
                fieldconst.Item = constvalue

                Dim isread As New IsEqualToType()
                isread.Item = findindex
                isread.FieldURIOrConstant = fieldconst

                Dim filter As New RestrictionType()
                filter.Item = isread

                'findItemRequest.Restriction = filter
                findItemRequest.ParentFolderIds = folderIDArray

                Dim fis As FindItemResponseType = esb.FindItem(findItemRequest)
                Dim rmta As ResponseMessageType() = fis.ResponseMessages.Items
                For Each rmt As ResponseMessageType In rmta
                    'Response.Write(rmt.MessageText + rmt.ResponseCode.ToString() + "<br/>") 
                    Dim firmt As FindItemResponseMessageType = DirectCast(rmt, FindItemResponseMessageType)
                    If firmt.ResponseClass = ResponseClassType.Success Then
                        Dim fit As FindItemParentType = firmt.RootFolder
                        Dim obj As Object = fit.Item

                        If TypeOf obj Is ArrayOfRealItemsType Then
                            Dim msgitems As ArrayOfRealItemsType = DirectCast(obj, ArrayOfRealItemsType)
                            If obj Is Nothing OrElse msgitems Is Nothing OrElse msgitems.Items Is Nothing Then
                                Exit For
                            End If
                            For iMsg As Integer = 0 To msgitems.Items.Length - 1
                                If TypeOf msgitems.Items(iMsg) Is MessageType Then

                                    'check if the message has been not read 
                                    Dim omsg As MessageType = TryCast(msgitems.Items(iMsg), MessageType)
                                    If Not omsg.IsRead Then
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
                                                        If msg.Sender IsNot Nothing AndAlso msg.Sender.Item IsNot Nothing Then
                                                            Dim sEmailFrom As EmailAddressType = msg.Sender.Item
                                                            If sEmailFrom.EmailAddress IsNot Nothing Then
                                                                'Dim sFrom As String = sEmailFrom.EmailAddress.ToLower()
                                                                'MsgBox(msg.Body.Value)
                                                                If IsSysFailEmail(msg) Then
                                                                    Dim nl As ArrayList = GetEmailLinkFromMsgBody(msg)
                                                                    If nl.Count > 0 Then
                                                                        Console.WriteLine("Subject : " + msg.Subject)
                                                                        Dim strReason As String = msg.Body.Value, strSubject As String = ""
                                                                        If msg.Subject IsNot Nothing AndAlso msg.Subject.Length > 0 Then
                                                                            strSubject = msg.Subject
                                                                            If strSubject.Length > 1000 Then
                                                                                strSubject = strSubject.Substring(0, 1000)
                                                                            End If
                                                                        End If
                                                                        If strReason.Length > 2000 Then
                                                                            strReason = strReason.Substring(0, 2000)
                                                                        End If
                                                                        'ws.AddInvalidEmailWithReason(nl.Item(0).ToString.Trim(), strSubject, strReason)
                                                                        Dim row As DataRow = dt.NewRow()
                                                                        With row
                                                                            .Item("EMAIL") = nl.Item(0).ToString.Trim() : .Item("INS_DATE") = Now : .Item("SUBJECT") = strSubject
                                                                            .Item("EMAIL_BODY") = strReason : .Item("REASON_FLAG") = "N/A"
                                                                            Dim bounce As BounceCode = GetBounceCodeAndMessage(strReason)
                                                                            If bounce IsNot Nothing Then
                                                                                If SoftBounceCode.Where(Function(x) bounce.Code.Contains(x) OrElse bounce.Code.Contains(x.Replace(".", ""))).Count > 0 Then
                                                                                    .Item("TYPE") = "SOFT"
                                                                                Else
                                                                                    .Item("TYPE") = "HARD"
                                                                                End If
                                                                                .Item("CODE") = bounce.Code.Trim
                                                                                .Item("MESSAGE") = bounce.Message.Trim
                                                                            End If
                                                                        End With
                                                                        dt.Rows.Add(row)
                                                                        Console.WriteLine("add invalid email:" + nl.Item(0).ToString.Trim())
                                                                        DeleteItem(esb, omsg.ItemId.Id)
                                                                    End If
                                                                Else
                                                                    If msg.Subject IsNot Nothing AndAlso _
                                                                    (msg.Subject.Contains("AutoReply") OrElse _
                                                                    msg.Subject.Contains("Out Of Office") OrElse _
                                                                    msg.Subject.Contains("Automatic Response")) Then
                                                                        DeleteItem(esb, omsg.ItemId.Id)
                                                                    Else
                                                                        Try
                                                                            If msg.From.Item.Name.Contains("postmaster@") AndAlso msg.Subject.Contains("Message delayed") Then
                                                                                DeleteItem(esb, omsg.ItemId.Id)
                                                                            Else
                                                                                DeleteItem(esb, omsg.ItemId.Id)
                                                                            End If
                                                                        Catch ex As Exception

                                                                        End Try
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Catch ex As Exception
                Dim BCopy As New SqlClient.SqlBulkCopy(strRFM)
                BCopy.DestinationTableName = "INVALID_EMAIL"
                BCopy.WriteToServer(dt)
                Console.WriteLine(ex.ToString)
                SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Sync Invalid Email error from Global", ex.ToString, True)
            End Try

            Dim BCopy1 As New SqlClient.SqlBulkCopy(strRFM)
            BCopy1.DestinationTableName = "INVALID_EMAIL"
            BCopy1.WriteToServer(dt)

            'Dim endws As New myadvan_global.EC()
            'endws.UseDefaultCredentials = True
            'endws.UpdateInvalidEmailReason()
            Exit Sub
            'Dim endws As New EC.EC()
            'endws.UseDefaultCredentials = True
            'endws.UpdateInvalidEmailReason()
            'Exit Sub
            Console.WriteLine("Sleeping...")
            Threading.Thread.Sleep(60 * 1000)
        End While
    End Sub
    Public Class BounceCode
        Public Code As String
        Public Message As String
    End Class
    Function GetBounceCodeAndMessage(ByVal strHtml As String) As BounceCode
        Dim code1 As New List(Of String), code2 As New List(Of String)
        For i As Integer = 4 To 9
            For j As Integer = 0 To 9
                For k As Integer = 0 To 9
                    code1.Add(i.ToString + j.ToString + k.ToString)
                    code2.Add(i.ToString + "." + j.ToString + "." + k.ToString)
                Next
            Next
        Next

        Dim bounce As New BounceCode
        Try
            Dim doc As New HtmlAgilityPack.HtmlDocument, targetIndex As Integer = -1
            doc.LoadHtml(strHtml)
            Dim pNodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//p")
            If pNodes IsNot Nothing AndAlso pNodes.Count > 0 Then
                For i As Integer = 0 To pNodes.Count - 1
                    If pNodes(i) IsNot Nothing AndAlso pNodes(i).InnerText.Contains("Generating server") Then
                        targetIndex = i + 1 : Exit For
                    End If
                Next

                strHtml = pNodes(targetIndex).InnerText
            Else
                targetIndex = 0
            End If

            If targetIndex >= 0 Then
                Dim rCode1 As String = code1.Where(Function(x) strHtml.Split(" ").Contains(x) OrElse strHtml.Split(" ").Contains("#" + x)).FirstOrDefault()
                Dim rCode2 As String = code2.Where(Function(x) strHtml.Split(" ").Contains(x)).FirstOrDefault()

                bounce.Code = rCode1 + " " + rCode2
                bounce.Message = strHtml
            End If

        Catch ex As Exception

        End Try

        Return bounce
    End Function
    Private Sub DeleteItem(ByVal esb As ExchangeServiceBinding, ByVal sItemId As String)
        ' Create the request. 
        Dim request As New DeleteItemType()

        ' Identify the items to delete. 
        Dim items As ItemIdType() = New ItemIdType(0) {}
        items(0) = New ItemIdType()
        items(0).Id = sItemId
        request.ItemIds = items

        ' Identify how deleted items are handled. 
        request.DeleteType = DisposalType.HardDelete

        ' Identify how tasks are deleted. 
        'request.AffectedTaskOccurrences = AffectedTaskOccurrencesType.SpecifiedOccurrenceOnly; 
        'request.AffectedTaskOccurrencesSpecified = true; 

        '''/ Identify how meeting cancellations are handled. 
        'request.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendOnlyToAll; 
        'request.SendMeetingCancellationsSpecified = true; 

        Try
            ' Send the response and receive the request. 
            Dim response As DeleteItemResponseType = esb.DeleteItem(request)
            Dim aormt As ArrayOfResponseMessagesType = response.ResponseMessages
            Dim rmta As ResponseMessageType() = aormt.Items

            ' Check each response message. 
            'foreach (ResponseMessageType rmt in rmta) 
            '{ 
            ' if (rmt.ResponseClass == ResponseClassType.Success) 
            ' { 
            ' //Console.WriteLine("Deleted item."); 
            ' } 
            '} 
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

        'specify itemshape 
        Dim itemproperties As New ItemResponseShapeType()
        itemproperties.BaseShape = DefaultShapeNamesType.AllProperties
        itemproperties.BodyType = BodyTypeResponseType.HTML
        itemproperties.BodyTypeSpecified = True
        itemproperties.IncludeMimeContent = True
        getrequest.ItemShape = itemproperties

        Dim itemtosend As ItemIdType() = New ItemIdType(0) {}
        itemtosend(0) = New ItemIdType()
        itemtosend(0).Id = sItemId

        getrequest.ItemIds = itemtosend

        Return getrequest
    End Function

    Public Function IsSysFailEmail(ByVal msg As MessageType) As Boolean
        If msg IsNot Nothing AndAlso msg.Subject IsNot Nothing AndAlso _
        (msg.Subject.StartsWith("未傳遞的主旨:") Or msg.Subject.StartsWith("Undeliverable:")) Then
            If msg.Body IsNot Nothing AndAlso msg.Body.Value IsNot Nothing AndAlso _
            (msg.Body.Value.Contains("傳送至下列收件者或通訊群組清單失敗:") OrElse _
             msg.Body.Value.Contains("Delivery has failed to these recipients or distribution lists")) Then
                Return True
            End If
        End If
        Return False
    End Function
    Public Function GetEmailLinkFromMsgBody(ByVal msg As MessageType) As ArrayList
        Dim al As New ArrayList
        If msg.Body Is Nothing OrElse msg.Body.Value Is Nothing Then Return al
        Dim doc As New HtmlAgilityPack.HtmlDocument
        doc.LoadHtml(msg.Body.Value)
        Dim nc As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//a")
        If Not IsDBNull(nc) AndAlso Not IsNothing(nc) Then
            For Each c As HtmlAgilityPack.HtmlNode In nc
                If c.HasAttributes AndAlso c.Attributes("href") IsNot Nothing AndAlso c.Attributes("href").Value.StartsWith("mailto:") Then
                    al.Add(c.Attributes("href").Value.Replace("mailto:", ""))
                End If
            Next
        End If
        Return al
    End Function

    Public Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal Subject As String, ByVal Body As String, ByVal IsBodyHtml As Boolean)
        Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
        htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
        htmlMessage.IsBodyHtml = IsBodyHtml
        htmlMessage.CC.Add("MyAdvantech@advantech.com")
        mySmtpClient = New Net.Mail.SmtpClient("172.21.34.21")
        mySmtpClient.Send(htmlMessage)
    End Sub

#Region "SendGrid"
    Public Enum BounceType
        Bounce
        Block
        Invalid
    End Enum
    Public Class SendGridBlockBounceRow
        Public Property status As String : Public Property created As Date : Public Property reason As String : Public Property email As String
    End Class
    Public Class SendGridDeleteBounceReturnMessage
        Public Property message As String
    End Class
    Public Function GetSendGridBounceList(ByVal BType As BounceType, ByVal SendGridID As String, ByVal SendGridPassword As String) As List(Of SendGridBlockBounceRow)
        Try
            Dim apiAction As String = ""
            Select Case BType
                Case BounceType.Block
                    apiAction = "blocks"
                Case BounceType.Bounce
                    apiAction = "bounces"
                Case BounceType.Invalid
                    apiAction = "invalidemails"
            End Select
            Dim wclient As New Net.WebClient()
            Dim JSserializer = New Script.Serialization.JavaScriptSerializer()
            Dim BlockResponse As String = wclient.DownloadString(String.Format("https://api.sendgrid.com/api/{0}.get.json?api_user={1}&api_key={2}&date=1", apiAction, SendGridID, SendGridPassword))
            Dim BlockResponseList As List(Of SendGridBlockBounceRow) = JSserializer.Deserialize(Of List(Of SendGridBlockBounceRow))(BlockResponse)
            Return BlockResponseList
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function DeleteSendGridBouncedList(ByVal BType As BounceType, ByVal Email As String, ByVal SendGridID As String, ByVal SendGridPassword As String) As Boolean
        Try
            Dim apiAction As String = ""
            Select Case BType
                Case BounceType.Block
                    apiAction = "blocks"
                Case BounceType.Bounce
                    apiAction = "bounces"
                Case BounceType.Invalid
                    apiAction = "invalidemails"
            End Select
            Dim wclient As New Net.WebClient()
            Dim JSserializer = New Script.Serialization.JavaScriptSerializer()
            Dim BlockResponse As String = wclient.DownloadString(String.Format("https://api.sendgrid.com/api/{0}.delete.json?api_user={1}&api_key={2}&email={3}", apiAction, SendGridID, SendGridPassword, Email))

            Dim response As SendGridDeleteBounceReturnMessage = JSserializer.Deserialize(Of SendGridDeleteBounceReturnMessage)(BlockResponse)
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Sub GetSendGridBounceMail(ByVal dtInvalid As DataTable, ByVal SendGridID As String, ByVal SendGridPassword As String)
        Try
            Dim conn As New SqlClient.SqlConnection(strMY)
            If conn.State <> ConnectionState.Open Then conn.Open()
            Dim cmd As SqlClient.SqlCommand = conn.CreateCommand() : cmd.CommandTimeout = 300 * 1000

            Dim Invalid As List(Of SendGridBlockBounceRow) = GetSendGridBounceList(BounceType.Invalid, SendGridID, SendGridPassword)
            If Invalid IsNot Nothing AndAlso Invalid.Count > 0 Then
                For Each item In Invalid
                    Dim counter As Integer = 0, ret As Boolean = False
                    While ret = False AndAlso counter < 3
                        ret = DeleteSendGridBouncedList(BounceType.Invalid, item.email, SendGridID, SendGridPassword)
                        counter += 1
                        Threading.Thread.Sleep(2 * 1000)
                    End While

                    If ret Then
                        Dim row As DataRow = dtInvalid.NewRow()
                        With row
                            .Item("EMAIL") = item.email : .Item("INS_DATE") = Now : .Item("SUBJECT") = "Undeliverable: " + GetMailSubject(conn, cmd, item.email)
                            .Item("EMAIL_BODY") = item.reason : .Item("REASON_FLAG") = "Invalid (by SendGrid)"
                            Dim bounce As BounceCode = GetBounceCodeAndMessage(item.reason)
                            If bounce IsNot Nothing Then
                                If SoftBounceCode.Where(Function(x) bounce.Code.Contains(x) OrElse bounce.Code.Contains(x.Replace(".", ""))).Count > 0 Then
                                    .Item("TYPE") = "SOFT"
                                Else
                                    .Item("TYPE") = "HARD"
                                End If
                                .Item("CODE") = bounce.Code.Trim
                                .Item("MESSAGE") = bounce.Message.Trim
                            End If
                        End With
                        dtInvalid.Rows.Add(row)
                        Console.WriteLine("SendGrid add invalid email:" + item.email)
                    End If

                Next
            End If

            Dim bounced As List(Of SendGridBlockBounceRow) = GetSendGridBounceList(BounceType.Bounce, SendGridID, SendGridPassword)
            If bounced IsNot Nothing AndAlso bounced.Count > 0 Then
                For Each item In bounced
                    Dim counter As Integer = 0, ret As Boolean = False
                    While ret = False AndAlso counter < 3
                        ret = DeleteSendGridBouncedList(BounceType.Bounce, item.email, SendGridID, SendGridPassword)
                        counter += 1
                        Threading.Thread.Sleep(2 * 1000)
                    End While

                    If ret Then
                        Dim row As DataRow = dtInvalid.NewRow()
                        With row
                            .Item("EMAIL") = item.email : .Item("INS_DATE") = Now : .Item("SUBJECT") = "Undeliverable: " + GetMailSubject(conn, cmd, item.email)
                            .Item("EMAIL_BODY") = item.reason : .Item("REASON_FLAG") = "Bounced (by SendGrid)"
                            Dim bounce As BounceCode = GetBounceCodeAndMessage(item.reason)
                            If bounce IsNot Nothing Then
                                If SoftBounceCode.Where(Function(x) bounce.Code.Contains(x) OrElse bounce.Code.Contains(x.Replace(".", ""))).Count > 0 Then
                                    .Item("TYPE") = "SOFT"
                                Else
                                    .Item("TYPE") = "HARD"
                                End If
                                .Item("CODE") = bounce.Code.Trim
                                .Item("MESSAGE") = bounce.Message.Trim
                            End If
                        End With
                        dtInvalid.Rows.Add(row)
                        Console.WriteLine("SendGrid add bounced email:" + item.email)
                    End If

                Next
            End If

            Dim Block As List(Of SendGridBlockBounceRow) = GetSendGridBounceList(BounceType.Block, SendGridID, SendGridPassword)
            If Block IsNot Nothing AndAlso Block.Count > 0 Then
                For Each item In Block
                    Dim counter As Integer = 0, ret As Boolean = False
                    While ret = False AndAlso counter < 3
                        ret = DeleteSendGridBouncedList(BounceType.Block, item.email, SendGridID, SendGridPassword)
                        counter += 1
                        Threading.Thread.Sleep(2 * 1000)
                    End While

                    If ret Then
                        Dim row As DataRow = dtInvalid.NewRow()
                        With row
                            .Item("EMAIL") = item.email : .Item("INS_DATE") = Now : .Item("SUBJECT") = "Undeliverable: " + GetMailSubject(conn, cmd, item.email)
                            .Item("EMAIL_BODY") = item.reason : .Item("REASON_FLAG") = "Blocked (by SendGrid)"
                            Dim bounce As BounceCode = GetBounceCodeAndMessage(item.reason)
                            If bounce IsNot Nothing Then
                                If SoftBounceCode.Where(Function(x) bounce.Code.Contains(x) OrElse bounce.Code.Contains(x.Replace(".", ""))).Count > 0 Then
                                    .Item("TYPE") = "SOFT"
                                Else
                                    .Item("TYPE") = "HARD"
                                End If
                                .Item("CODE") = bounce.Code.Trim
                                .Item("MESSAGE") = bounce.Message.Trim
                            End If
                        End With
                        dtInvalid.Rows.Add(row)
                        Console.WriteLine("SendGrid add blocked email:" + item.email)
                    End If

                Next
            End If
            Dim BCopy1 As New SqlClient.SqlBulkCopy(strRFM)
            BCopy1.DestinationTableName = "INVALID_EMAIL"
            BCopy1.WriteToServer(dtInvalid)
            conn.Close()
        Catch ex As Exception
            Dim BCopy1 As New SqlClient.SqlBulkCopy(strRFM)
            BCopy1.DestinationTableName = "INVALID_EMAIL"
            BCopy1.WriteToServer(dtInvalid)
            SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Sync Invalid Email error from SendGrid", ex.ToString, True)
        End Try
    End Sub

    Public Function GetMailSubject(ByVal conn As SqlClient.SqlConnection, ByVal cmd As SqlClient.SqlCommand, ByVal Email As String) As String
        If conn.State <> ConnectionState.Open Then conn.Open()
        cmd.CommandText = String.Format(" select isnull(b.EMAIL_SUBJECT,'') as EMAIL_SUBJECT from CAMPAIGN_SEND_LOG a with (NOLOCK) " + _
                                        " inner join CAMPAIGN_MASTER b with (NOLOCK) on a.CAMPAIGN_ROW_ID=b.ROW_ID " + _
                                        " where a.CONTACT_EMAIL='{0}' order by a.SEND_DATE desc ", Email)
        Dim obj As Object = cmd.ExecuteScalar()
        If obj IsNot Nothing Then Return obj.ToString
        Return ""
    End Function

#End Region

    'Partial Public Class ExchangeServiceBinding
    '    Inherits SoapHttpClientProtocol
    '    Protected Overloads Overrides Function GetReaderForMessage(ByVal message As SoapClientMessage, ByVal bufferSize As Integer) As XmlReader
    '        Dim retval As XmlReader = MyBase.GetReaderForMessage(message, bufferSize)
    '        Dim xrt As XmlTextReader = TryCast(retval, XmlTextReader)
    '        If xrt IsNot Nothing Then
    '            xrt.Normalization = False
    '        End If
    '        Return retval
    '    End Function
    'End Class

End Module
