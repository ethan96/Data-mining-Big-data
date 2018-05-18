Module Module1

    Sub Main()
        Dim ms As New System.Net.Mail.SmtpClient(CenterLibrary.AppConfig.SMTPServerIP)
        Try
            Dim ws As New MyAutoPrj.AutoJob
            ws.Timeout = -1
            If ws.SyncCatalogPriceATP() Then
                Exit Sub
                Dim dt As DataTable = GetDt()
                Dim custAry As New ArrayList
                For Each r As DataRow In dt.Rows
                    If custAry.Contains(r.Item("customer_email")) = False Then custAry.Add(r.Item("customer_email"))
                Next

                Dim sqlConn As New SqlClient.SqlConnection(CenterLibrary.DBConnection.MyAdvantechGlobal)
                sqlConn.Open()
                For Each c As String In custAry
                    Dim rs() As DataRow = dt.Select("customer_email='" + c + "'")
                    If rs.Length > 0 Then
                        Dim o As Object = Nothing, cname As String = "", bccAry As New ArrayList
                        For Each r As DataRow In rs
                            If bccAry.Contains(r.Item("owner_email")) = False Then bccAry.Add(r.Item("owner_email"))
                        Next
                        Dim cmd As New SqlClient.SqlCommand("select top 1 firstname from siebel_contact where email_address='" + c + "'")
                        cmd.Connection = sqlConn
                        o = cmd.ExecuteScalar()
                        If o IsNot Nothing Then
                            cname = o.ToString()
                        Else
                            cname = Split(c, "@")(0)
                        End If
                        Dim mbody As String = Notify(cname, rs)
                        c = "tc.chen@advantech.com.tw"
                        Dim m As New System.Net.Mail.MailMessage("MyAdvantech@advantech.com", c, "The Catalog Forecast you placed has arrived", mbody)
                        m.Bcc.Add("chentc@gmail.com") : m.Bcc.Add("chen.tc@gmail.com") : m.IsBodyHtml = True
                        ms.Send(m)
                    End If
                Next
                sqlConn.Close()
            End If
            ms.Send("myadvantech@advantech.com", "MyAdvantech@advantech.com", "sync CatalogPriceATP ok", "")
        Catch ex As Exception
            ms.Send("myadvantech@advantech.com", "MyAdvantech@advantech.com", "sync CatalogPriceATP Error", ex.ToString())
        End Try
    End Sub

    Function GetDt() As DataTable
        Dim sb As New System.Text.StringBuilder
        With sb
            .AppendLine(String.Format(" select a.PART_NO, a.atp, a.Price, z1.user_id as customer_email, IsNull(z3.EXTENDED_DESC, z1.description) as product_desc, "))
            .AppendLine(String.Format(" z1.qty as request_qty, z2.owner_email, z1.date as request_date "))
            .AppendLine(String.Format(" from CATALOG_PRICE_ATP a inner join [aclsql6\sql2008r2].MyAdvantechGlobal.dbo.forecast_catalog_history_new z1 on a.PART_NO=z1.part_no "))
            .AppendLine(String.Format(" inner join [aclsql6\sql2008r2].MyAdvantechGlobal.dbo.FORECAST_CATALOG_LIST z2 on z1.catalog_id=z2.row_id "))
            .AppendLine(String.Format(" left join mylocal.dbo.sap_product_ext_desc z3 on a.PART_NO=z3.part_no "))
            .AppendLine(String.Format(" where a.ATP>0 and z1.IS_INFORMED=0 and z1.user_id like '%@%.%' "))
        End With
        Dim apt As New SqlClient.SqlDataAdapter(sb.ToString(), System.Configuration.ConfigurationManager.ConnectionStrings("MyDM").ConnectionString)
        Dim dt As New DataTable
        apt.Fill(dt)
        Return dt
    End Function

    Function Notify(ByVal custName As String, ByRef rs As DataRow()) As String
        Dim sb As New System.Text.StringBuilder
        With sb
            .AppendLine(String.Format("Dear Mr/Mrs {0},<br />", custName))
            .AppendLine(String.Format("<br />"))
            .AppendLine(String.Format("This is a notice to remind you that the catalog you've requested with forecast has arrived.<br />"))
            .AppendLine(String.Format("They are now available to be ordered with your products:<br />"))
            .AppendLine(String.Format("<br />"))
            .AppendLine(String.Format("<table style='border-width:1px; border-style:groove'>"))
            .AppendLine(String.Format("<tr><th>Part#</th><th>Theme</th><th>Your forecast</th><th>Qty available</th><th>Cost/USD</th></tr>"))
            For Each r As DataRow In rs
                .AppendLine(String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>", _
                                          r.Item("PART_NO"), r.Item("product_desc"), r.Item("request_qty").ToString(), r.Item("atp").ToString(), r.Item("Price").ToString()))
            Next
            .AppendLine(String.Format("</table>"))
            .AppendLine(String.Format("<br />"))
            .AppendLine(String.Format("For better timing of utility, please request them ASAP. <br />"))
            .AppendLine(String.Format("In case the catalogs' owner is not identified for 3 months, the system will automatically allocate to other departments who needs.<br />"))
            .AppendLine(String.Format("<br />"))
            .AppendLine(String.Format("Thanks for your attention and if you need any help, please feel free to contact us via inquiry@advantech.com<br />"))
        End With
        Return sb.ToString()
    End Function

End Module
