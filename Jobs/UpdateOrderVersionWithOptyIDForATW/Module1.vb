Imports System.Text
Module Module1

    Sub Main()
        Dim dt1 As DateTime = DateTime.Now
        Dim dt1_First As String = dt1.AddDays(-1).ToString("yyyyMMdd")
        Dim sb As New StringBuilder
        sb.AppendFormat(" select a.vbeln, a.erdat, a.VSNMR_V as ver ")
        sb.AppendFormat(" from saprdp.vbak a  ")
        sb.AppendFormat(" where a.mandt='168' and a.vbeln like 'TWO%' and a.vkorg='TW01' and a.erdat>='{0}' ", dt1_First)
        sb.AppendFormat(" and a.VSNMR_V in ('NEW ID',' ')  ")
        sb.AppendFormat(" order by a.erdat desc ")
        Dim dt As DataTable = dbUtil.dbOracleGetDataTable("SAP_PRD", sb.ToString())
        Dim ordernos As String() = dt.AsEnumerable().Select(Function(d) d.Field(Of String)("vbeln")).ToArray()
        Dim keystr As String = String.Join("','", ordernos).Trim(",'")
        sb.Clear()
        sb.AppendFormat(" select  A.cartid,A.OpportunityID,B.OrderNo   from  [CARTMASTERV2]  A  ")
        sb.AppendFormat(" INNER JOIN [Cart2OrderMaping] B ON B.CARTID = A.CartID ")
        sb.AppendFormat(" where   B.OrderNo in ('{0}') AND A.OpportunityID <>'' AND A.OpportunityID IS NOT NULL", keystr)
        dt.Clear()
        dt = CenterLibrary.dbUtil.dbGetDataTable(CenterLibrary.DBConnection.MyAdvantechGlobal, sb.ToString())
        Dim ReturnTable As New DataTable()
        Dim drs As IEnumerable(Of DataRow) = dt.AsEnumerable()
        Dim i As Integer = 0
        For Each dr As DataRow In drs
            Dim retbool As Boolean = SAPDAL.UpdateSOZeroPriceItems(dr.Item("OrderNo"), dr.Item("OpportunityID"), ReturnTable)
            If Not retbool Then
                i = i + 1
                MailUtil.SendEmail("ming.zhao@advantech.com.cn", "ming.zhao@advantech.com.cn", "call Change_SD_Order failed", dr.Item("OrderNo") + ":" + dr.Item("OpportunityID"), True)
            End If
            Console.WriteLine(dr.Item("OrderNo") + ":" + dr.Item("OpportunityID") + "--" + retbool.ToString())
            If i > 4 Then
                Exit For
            End If
        Next
    End Sub

End Module
