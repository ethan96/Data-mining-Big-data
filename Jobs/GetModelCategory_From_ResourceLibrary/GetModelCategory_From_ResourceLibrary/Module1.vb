Imports System.Text.RegularExpressions

Module Module1

    Public strMY As String = CenterLibrary.DBConnection.MyAdvantechGlobal
    Sub Main()
        Dim intApiWs As New CurationWs.IntApi
        Dim ConnMY As New SqlClient.SqlConnection(strMY)
        Dim dtCategory As New DataTable
        Try
            Dim aptCategory As New SqlClient.SqlDataAdapter("select distinct t.ITEM from ( " + _
            " select distinct ITEM_DISPLAYNAME as ITEM from PIS.dbo.MODELCATEGORY_INTERESTEDPRODUCT_MAPPING " + _
            " union all " + _
            " select distinct INTERESTED_PRODUCT_DISPLAY_NAME as ITEM from PIS.dbo.MODELCATEGORY_INTERESTEDPRODUCT_MAPPING " + _
            " union all " + _
            " select distinct PRODUCT_GROUP_DISPLAY_NAME as ITEM from PIS.dbo.MODELCATEGORY_INTERESTEDPRODUCT_MAPPING " + _
            " union all " + _
            " select distinct MODEL_NO as ITEM FROM SAP_PRODUCT where MATERIAL_GROUP='PRODUCT' " + _
            " union all " + _
            " select distinct PART_NO as ITEM FROM SAP_PRODUCT where MATERIAL_GROUP='PRODUCT' " + _
            " ) as t where t.ITEM not in ('','-',' ','00','R1','ESS','　19C6969691','(DEL)9680001337','1','security','API','BIOS','OS','Memory','CPU','GPS','Stand','Multifunction','counter','HDD','Accessory','power supply','Bluetooth','UPS','WiFi') " + _
            " order by t.ITEM ", ConnMY)
            aptCategory.Fill(dtCategory)
        Catch ex As Exception
            CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Get Resource Failed: Get All Category error", ex.ToString, True, "", "")
        End Try
        
        'Dim arrCategory As New ArrayList
        'For Each row As DataRow In dtCategory.Rows
        '    arrCategory.Add(Replace(row.Item("ITEM"), "'", "''"))
        'Next
        'Dim strCategorys As String = String.Join("|", arrCategory.ToArray())
        'Dim RegExp As New Regex(strCategorys, RegexOptions.IgnoreCase)

        Dim strSelectResourceSql As String = _
            " select distinct a.DOC_ID, 'SUPPORT' as DOC_SOURCE,isnull(a.DESCRIPTION,'') + isnull(a.SEARCH_KEY,'') + isnull(a.ABSTRACT,'') + ISNULL(cast(c.RESOLUTION_TEXT as nvarchar),'') as DOC_DESC from SUPPORT_DOWNLOAD a left join SIEBEL_SR_SOLUTION_RELATION b on a.DOC_ID=b.SR_ID left join SIEBEL_SR_SOLUTION c on b.SOLUTION_ID=c.SR_ID where a.ISSUE_DATE between DATEADD(day,-7,convert(varchar(10),getdate(),111)) and DATEADD(day,1,convert(varchar(10),getdate(),111)) or c.LAST_UPD between DATEADD(day,-7,convert(varchar(10),getdate(),111)) and DATEADD(day,1,convert(varchar(10),getdate(),111)) " + _
            " union all " + _
            " select distinct a.LITERATURE_ID as DOC_ID, 'PIS' as DOC_SOURCE, isnull((select z.model_name from PIS.dbo.Model_lit z where z.literature_id=a.LITERATURE_ID for xml path('')),'') + ISNULL(a.LIT_NAME,'') + ISNULL(a.LIT_DESC,'') + ISNULL(a.FILE_NAME,'') + ISNULL(b.TXT_CONTENT,'') as DOC_DESC from PIS.dbo.LITERATURE a left join SIEBEL_LITERATURE_DETAIL b on a.LITERATURE_ID=b.LIT_ID where a.CREATED between DATEADD(day,-7,convert(varchar(10),getdate(),111)) and DATEADD(day,1,convert(varchar(10),getdate(),111)) or a.LAST_UPDATED between DATEADD(day,-7,convert(varchar(10),getdate(),111)) and DATEADD(day,1,convert(varchar(10),getdate(),111)) " + _
            " union all " + _
            " select distinct a.RECORD_ID as DOC_ID, 'CMS' as DOC_SOURCE, ISNULL(a.TITLE,'') + ISNULL(a.ABSTRACT,'') + ISNULL(b.CMS_CONTENT,'') as DOC_DESC from WWW_RESOURCES a left join WWW_RESOURCES_DETAIL b on a.RECORD_ID=b.RECORD_ID where a.RELEASE_DATE between DATEADD(day,-7,convert(varchar(10),getdate(),111)) and DATEADD(day,1,convert(varchar(10),getdate(),111)) or a.LASTUPDATED between DATEADD(day,-7,convert(varchar(10),getdate(),111)) and DATEADD(day,1,convert(varchar(10),getdate(),111)) " + _
            " union all " + _
            " select a.ROW_ID as DOC_ID, 'EDM' as DOC_SOURCE, isnull((select z.MODEL_NO from CAMPAIGN_MODEL_CATEGORY z where z.CAMPAIGN_ROW_ID=a.ROW_ID for XML path('')),'')+a.CAMPAIGN_NAME+ isnull(a.DESCRIPTION,'')+ISNULL(a.TEMPLATE_FILE_TEXT,'') as DOC_DESC from CAMPAIGN_MASTER a where a.IS_DISABLED=0 and a.ACTUAL_SEND_DATE is not null and a.ACTUAL_SEND_DATE between DATEADD(day,-7,convert(varchar(10),getdate(),111)) and DATEADD(day,1,convert(varchar(10),getdate(),111)) "
        'Dim strSelectResourceSql As String = _
        '    " select distinct a.DOC_ID, 'SUPPORT' as DOC_SOURCE,isnull(a.DESCRIPTION,'') + isnull(a.SEARCH_KEY,'') + isnull(a.ABSTRACT,'') + ISNULL(cast(c.RESOLUTION_TEXT as nvarchar),'') as DOC_DESC from SUPPORT_DOWNLOAD a left join SIEBEL_SR_SOLUTION_RELATION b on a.DOC_ID=b.SR_ID left join SIEBEL_SR_SOLUTION c on b.SOLUTION_ID=c.SR_ID  " + _
        '    " union all " + _
        '    " select distinct a.LITERATURE_ID as DOC_ID, 'PIS' as DOC_SOURCE, isnull((select z.model_name from PIS.dbo.Model_lit z where z.literature_id=a.LITERATURE_ID for xml path('')),'') + ISNULL(a.LIT_NAME,'') + ISNULL(a.LIT_DESC,'') + ISNULL(a.FILE_NAME,'') + ISNULL(b.TXT_CONTENT,'') as DOC_DESC from PIS.dbo.LITERATURE a left join SIEBEL_LITERATURE_DETAIL b on a.LITERATURE_ID=b.LIT_ID  " + _
        '    " union all " + _
        '    " select distinct a.RECORD_ID as DOC_ID, 'CMS' as DOC_SOURCE, ISNULL(a.TITLE,'') + ISNULL(a.ABSTRACT,'') + ISNULL(b.CMS_CONTENT,'') as DOC_DESC from WWW_RESOURCES a left join WWW_RESOURCES_DETAIL b on a.RECORD_ID=b.RECORD_ID " + _
        '    " union all " + _
        '    " select a.ROW_ID as DOC_ID, 'EDM' as DOC_SOURCE, isnull((select z.MODEL_NO from CAMPAIGN_MODEL_CATEGORY z where z.CAMPAIGN_ROW_ID=a.ROW_ID for XML path('')),'')+a.CAMPAIGN_NAME+ isnull(a.DESCRIPTION,'')+ISNULL(a.TEMPLATE_FILE_TEXT,'') as DOC_DESC from CAMPAIGN_MASTER a where a.IS_DISABLED=0 and a.ACTUAL_SEND_DATE is not null  "

        Dim dtResource As New DataTable
        Try
            Dim aptResource As New SqlClient.SqlDataAdapter(strSelectResourceSql, ConnMY)
            aptResource.Fill(dtResource)
        Catch ex As Exception
            CenterLibrary.MailUtil.SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Get Resource Failed: Get All Resources error", ex.ToString, True, "", "")
        End Try

        'Try
        For Each row As DataRow In dtResource.Rows
            Dim arrMatchedCategorys() As String = {}
            Dim strDocDesc As String = row.Item("DOC_DESC")
            If Not String.IsNullOrEmpty(strDocDesc) AndAlso strDocDesc <> "" Then
                Dim mmAry As New List(Of String), base As Integer = 0
                While dtCategory.Rows.Count - (1000 * base) >= 0
                    Dim arrCategory As New ArrayList
                    For i As Integer = 1000 * base To (1000 * (base + 1)) - 1
                        If i <= dtCategory.Rows.Count - 1 Then
                            arrCategory.Add(Replace(dtCategory.Rows(i).Item("ITEM").ToString.Trim, "'", "''"))
                        Else
                            Exit For
                        End If
                    Next
                    Dim strCategorys As String = String.Join("|", arrCategory.ToArray())
                    Dim RegExp As New Regex(strCategorys, RegexOptions.IgnorePatternWhitespace)
                    Dim mc As MatchCollection = RegExp.Matches(strDocDesc)
                    For Each m As Match In mc
                        If m.Value.Trim <> "" Then If Not mmAry.Contains("'" + m.Value + "'") Then mmAry.Add("'" + m.Value + "'")
                    Next
                    base += 1
                End While
                arrMatchedCategorys = mmAry.ToArray()
            End If
            Try
                Dim cmd As New SqlClient.SqlCommand("delete from RESOURCE_MODEL_CATEGORY where DOC_ID=@DOCID", ConnMY)
                cmd.Parameters.AddWithValue("DOCID", row.Item("DOC_ID"))
                If ConnMY.State <> ConnectionState.Open Then ConnMY.Open()
                cmd.ExecuteNonQuery()
                If arrMatchedCategorys IsNot Nothing AndAlso arrMatchedCategorys.Count > 0 Then
                    Dim dtPDGroup As New DataTable
                    Dim aptPDGroup As New SqlClient.SqlDataAdapter(String.Format("select distinct INTERESTED_PRODUCT_CATEGOEY_ID, PRODUCT_GROUP_CATEGOEY_ID from PIS.dbo.MODELCATEGORY_INTERESTEDPRODUCT_MAPPING where ITEM_DISPLAYNAME in ({0}) or INTERESTED_PRODUCT_DISPLAY_NAME in ({0}) or PRODUCT_GROUP_DISPLAY_NAME in ({0})", String.Join(",", arrMatchedCategorys)), ConnMY)
                    aptPDGroup.Fill(dtPDGroup)
                    For Each rowPD As DataRow In dtPDGroup.Rows
                        cmd = New SqlClient.SqlCommand(" insert into RESOURCE_MODEL_CATEGORY (DOC_ID,DOC_SOURCE,MATCH_KEY,MATCH_INTPROD,MATCH_PDGROUP) " + _
                                                       " values (@DOC_ID,@DOC_SOURCE,@MATCH_KEY,@MATCH_INTPROD,@MATCH_PDGROUP) ", ConnMY)
                        cmd.Parameters.AddWithValue("DOC_ID", row.Item("DOC_ID")) : cmd.Parameters.AddWithValue("DOC_SOURCE", row.Item("DOC_SOURCE"))
                        cmd.Parameters.AddWithValue("MATCH_KEY", String.Join(",", arrMatchedCategorys))
                        cmd.Parameters.AddWithValue("MATCH_INTPROD", rowPD.Item("INTERESTED_PRODUCT_CATEGOEY_ID"))
                        cmd.Parameters.AddWithValue("MATCH_PDGROUP", rowPD.Item("PRODUCT_GROUP_CATEGOEY_ID"))
                        If ConnMY.State <> ConnectionState.Open Then ConnMY.Open()
                        cmd.ExecuteNonQuery()
                    Next
                    Console.WriteLine("DOC_ID: " + row.Item("DOC_ID") + " for " + String.Join(",", arrMatchedCategorys))
                End If
            Catch ex As Exception
                Console.WriteLine("Failed: DOC_ID: " + row.Item("DOC_ID") + " for " + String.Join(",", arrMatchedCategorys))
            End Try
            ConnMY.Close()
        Next
        'Catch ex As Exception
        '    SendEmail("rudy.wang@advantech.com.tw", "MyAdvantech@advantech.com", "Error", ex.ToString, True)
        'End Try

        Console.WriteLine("OK")
    End Sub

    'Public Sub SendEmail(ByVal SendTo As String, ByVal From As String, ByVal Subject As String, ByVal Body As String, ByVal IsBodyHtml As Boolean)
    '    Dim htmlMessage As Net.Mail.MailMessage, mySmtpClient As Net.Mail.SmtpClient
    '    htmlMessage = New Net.Mail.MailMessage(From, SendTo, Subject, Body)
    '    htmlMessage.IsBodyHtml = IsBodyHtml
    '    mySmtpClient = New Net.Mail.SmtpClient("172.20.0.76")
    '    mySmtpClient.Send(htmlMessage)
    'End Sub
End Module
