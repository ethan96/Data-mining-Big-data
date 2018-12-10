Module Module1
    Public _strACLSql5 As String = "Data Source=aclsql6\sql2008r2;Initial Catalog=MyAdvantechGlobal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"
    Public _strMyGlob As String = "Data Source=myadvan-global;Initial Catalog=MyLocal;Persist Security Info=True;User ID=b2bsa;Password=@dvantech!;async=true;Connect Timeout=300;pooling='true'"

    Sub Main()
        Dim aclsql5Conn As New SqlClient.SqlConnection(_strACLSql5), myglobConn As New SqlClient.SqlConnection(_strMyGlob)
        Dim pdt As New DataTable
        Dim da As New SqlClient.SqlDataAdapter( _
            "  select top 9999 PART_NO " + _
            "  from sap_product  " + _
            "  where model_no<>'' and left(part_no,1) not in ('#','$','1','2','9','Y') and LEFT(part_no,3) not in ('NRE')  " + _
            "  and material_group in ('PRODUCT','BTOS') and STATUS in ('A','N','H') " + _
            "  order by part_no ", aclsql5Conn)
        da.SelectCommand.CommandTimeout = 5 * 60
        Try
            da.Fill(pdt)
        Catch ex As Exception
            aclsql5Conn.Close()
            Console.WriteLine(ex.ToString()) : Console.Read() : Exit Sub
        End Try
        myglobConn.Open()
        Dim bk As New SqlClient.SqlBulkCopy(_strMyGlob)
        bk.DestinationTableName = "DM_BASKET_ANALYSIS"
        For Each pr As DataRow In pdt.Rows
            Dim pn As String = pr.Item("part_no"), qdt As New DataTable
            Dim strQ As String = String.Format( _
                " select top 30 '{0}' as PART_NO, a.item_no as ref_part_no, count(a.order_no) as orders, cast(sum(a.us_amt) as float) as US_AMOUNT " + _
                "  from eai_sale_fact a   " + _
                "  where a.sector in ('DMF','AOnline','IA-AOnline','EP-AOnline') and a.factyear>=2010 " + _
                "  and a.order_no in   " + _
                "  (  " + _
                "  	select distinct b.order_no   " + _
                "  	from eai_sale_fact b   " + _
                "  	where b.item_no='{0}' and b.sector in ('DMF','AOnline','IA-AOnline','EP-AOnline') and b.factyear>=2010  " + _
                "  	and a.item_no<>'{0}'  " + _
                "  ) " + _
                "  group by a.item_no   " + _
                "  order by count(a.order_no) desc  ", pn)
            Console.WriteLine("proc " + pn)
            da = New SqlClient.SqlDataAdapter(strQ, aclsql5Conn)
            da.SelectCommand.CommandTimeout = 5 * 60
            Try
                da.Fill(qdt)
            Catch ex As Exception
                aclsql5Conn.Close()
                Console.WriteLine(ex.ToString()) : Console.Read() : Exit Sub
            End Try
            If qdt.Rows.Count > 0 Then
                Try
                    Dim _cmd As New SqlClient.SqlCommand("delete from DM_BASKET_ANALYSIS where part_no='" + pn + "'", myglobConn)
                Catch ex As Exception
                    myglobConn.Close()
                    Console.WriteLine(ex.ToString()) : Console.Read() : Exit Sub
                End Try
                Try
                    bk.WriteToServer(qdt)
                    Console.WriteLine("wrote " + qdt.Rows.Count.ToString() + " rows of " + pn)
                Catch ex As Exception
                    myglobConn.Close()
                    Console.WriteLine(ex.ToString()) : Console.Read() : Exit Sub
                End Try

            End If
        Next
        myglobConn.Close() : aclsql5Conn.Close()
        Console.WriteLine("done")
        'Console.Read()
    End Sub



    Public Function dbGetDataTable(ByVal conn As String, ByVal strSqlCmd As String) As DataTable
        Dim g_adoConn As New SqlClient.SqlConnection(conn)
        Dim dt As New DataTable
        Dim da As New SqlClient.SqlDataAdapter(strSqlCmd, g_adoConn)
        da.SelectCommand.CommandTimeout = 5 * 60
        da.Fill(dt)
        g_adoConn.Close() : g_adoConn = Nothing
        Return dt
    End Function

End Module
