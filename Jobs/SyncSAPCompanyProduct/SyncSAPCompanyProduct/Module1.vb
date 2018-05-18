Module Module1

    Sub Main(ByVal args() As String)
        Dim qDate As Date
        Dim tw As New System.Globalization.CultureInfo("zh-TW")
        If Not args Is Nothing AndAlso args.Length = 1 AndAlso Date.TryParseExact(args(0), "yyyyMMdd", tw, System.Globalization.DateTimeStyles.None, qDate) Then
            'If date format is true, then qDate = args(0)
        Else
            qDate = DateAdd(DateInterval.Day, -1, Now)
        End If
        Dim ODtComp As New DataTable()
        Try
            OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP Company Start: " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "", False)
            ODtComp = OraDbUtil.dbGetDataTable("SAP_PRD",
                " SELECT distinct a.objectid as company_id " &
                " FROM saprdp.CDHDR a " &
                " WHERE a.mandant='168' and a.OBJECTCLAS = 'DEBI' and a.UDATE>='" & qDate.ToString("yyyyMMdd") & "' AND (a.TCODE = 'VD01' or a.TCODE = 'VD02'or a.TCODE = 'XD01'or a.TCODE = 'XD02') " &
                " order by a.objectid")
        Catch ex As Exception
            OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP Company Error: " + Now.ToString("yyyy-MM-dd HH:mm:ss"), ex.Message, False)
        End Try
        If Not IsNothing(ODtComp) AndAlso ODtComp.Rows.Count > 0 Then
            Dim mbody As New System.Text.StringBuilder
            Try
                Dim arrcom As New ArrayList
                For i As Integer = 0 To ODtComp.Rows.Count - 1
                    arrcom.Add(ODtComp.Rows(i).Item("company_id"))
                    If i > 0 AndAlso i Mod 20 = 0 Then
                        Dim subem As String = ""
                        Dim fmsg As New ArrayList
                        SAPDAL.syncSingleCompany.syncSingleSAPCustomer(arrcom, False, subem, fmsg)
                        arrcom.Clear()
                        If subem <> "" Then
                            mbody.AppendLine(subem)
                        End If
                        If Not IsNothing(fmsg) AndAlso fmsg.Count > 0 Then
                            For Each r As String In fmsg
                                mbody.AppendLine(r)
                            Next
                        End If
                    ElseIf i > 0 AndAlso i = ODtComp.Rows.Count - 1 Then
                        Dim subem As String = ""
                        Dim fmsg As New ArrayList
                        SAPDAL.syncSingleCompany.syncSingleSAPCustomer(arrcom, False, subem, fmsg)
                        arrcom.Clear()
                        If subem <> "" Then
                            mbody.AppendLine(subem)
                        End If
                        If Not IsNothing(fmsg) AndAlso fmsg.Count > 0 Then
                            For Each r As String In fmsg
                                mbody.AppendLine(r)
                            Next
                        End If
                    End If
                Next
                'ICC 2014/10/01 Truncate SAP_COMPANY_CONTACTS and sync from SAP
                SAPDAL.syncSingleCompany.syncSapCompanyContacts()
                'ICC 2014/12/18 Truncate SAP_COMPANY_LOV and sync from SAP
                SAPDAL.syncSingleCompany.syncSapCompanyLov()
                'ICC 2017/8/25 Sync SAP_DIMCOMPANY_EXT
                SAPDAL.syncSingleCompany.syncSAPCompnayEXT()
            Catch ex As Exception
                mbody.AppendLine(ex.Message.ToString)
            Finally
                OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP Company End: " + Now.ToString("yyyy-MM-dd HH:mm:ss"), mbody.ToString, False)
            End Try
            Console.WriteLine(mbody)
        Else
            OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP Company End: " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "No data! ", False)
        End If

        Dim ODTProd As New DataTable()
        Try
            OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP Product Start: " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "", False)
            ODTProd = OraDbUtil.dbGetDataTable("SAP_PRD",
                                                                " SELECT distinct a.objectid as part_no " &
                                                                " FROM saprdp.CDHDR a " &
                                                                " WHERE a.mandant='168' and a.OBJECTCLAS = 'MATERIAL' and a.UDATE>='" & qDate.ToString("yyyyMMdd") & "' and (a.TCODE LIKE 'MM01%' or a.TCODE LIKE 'MM02%') " &
                                                                " order by a.objectid")
        Catch ex As Exception
            OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP Product Error: " + Now.ToString("yyyy-MM-dd HH:mm:ss"), ex.Message, False)
        End Try
        If Not IsNothing(ODTProd) AndAlso ODTProd.Rows.Count > 0 Then
            Dim mbody As New System.Text.StringBuilder
            Try
                Dim arrprod As New ArrayList
                For i As Integer = 0 To ODTProd.Rows.Count - 1
                    arrprod.Add(ODTProd.Rows(i).Item("part_no"))
                    If i > 0 AndAlso i Mod 200 = 0 Then
                        Dim subem As String = ""
                        Dim fmsg As New ArrayList
                        SAPDAL.syncSingleProduct.syncSAPProduct(arrprod, "ALL", False, subem, True, fmsg, True)
                        arrprod.Clear()
                        If subem <> "" Then
                            mbody.AppendLine(subem)
                        End If
                        If Not IsNothing(fmsg) AndAlso fmsg.Count > 0 Then
                            For Each r As String In fmsg
                                mbody.AppendLine(r)
                            Next
                        End If
                    ElseIf i > 0 AndAlso i = ODTProd.Rows.Count - 1 Then
                        Dim subem As String = ""
                        Dim fmsg As New ArrayList
                        SAPDAL.syncSingleProduct.syncSAPProduct(arrprod, "ALL", False, subem, True, fmsg, True)
                        arrprod.Clear()
                        If subem <> "" Then
                            mbody.AppendLine(subem)
                        End If
                        If Not IsNothing(fmsg) AndAlso fmsg.Count > 0 Then
                            For Each r As String In fmsg
                                mbody.AppendLine(r)
                            Next
                        End If
                    End If
                Next
            Catch ex As Exception
                mbody.AppendLine(ex.Message)
            Finally
                OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP Product End: " + Now.ToString("yyyy-MM-dd HH:mm:ss"), mbody.ToString, False)
            End Try
            Console.WriteLine(mbody)
        Else
            OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP Product Error: " + Now.ToString("yyyy-MM-dd HH:mm:ss"), "Sync SAP product error! No data", False)
        End If
        updateExtDesc()

        If DateTime.Now.DayOfWeek = DayOfWeek.Sunday Then
            Dim result As String = SAPDAL.syncSingleProduct.SyncSAPProudctWarranty()
            OraDbUtil.SendEmail("myadvantech@advantech.com", "myadvantech@advantech.com", "Sync SAP product warranty month end", result.Replace(",", "<br />"), True)
        End If
    End Sub

    Public Sub updateExtDesc()
        Dim str As String = " update SAP_PRODUCT " & _
        " set SAP_PRODUCT.PRODUCT_DESC = b.extended_desc " & _
        " from SAP_PRODUCT a inner join SAP_PRODUCT_EXT_DESC b on a.part_no=b.PART_NO; " & _
        " update cbom_catalog_category " & _
        " set cbom_catalog_category.CATEGORY_DESC = b.extended_desc " & _
        " from cbom_catalog_category a inner join SAP_PRODUCT_EXT_DESC b on a.CATEGORY_ID=b.PART_NO " & _
        " where a.ORG<>'EU'"

        SAPDAL.dbUtil.dbExecuteNoQuery("MY", str)

    End Sub
End Module
