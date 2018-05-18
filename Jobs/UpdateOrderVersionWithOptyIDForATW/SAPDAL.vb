
Public Class SAPDAL
    Public Shared Function UpdateSOZeroPriceItems(ByVal SO_NO As String, ByVal optyID As String, ByRef ReturnTable As DataTable) As Boolean
        'Dim aptOrderDetail As New MyOrderDSTableAdapters.ORDER_DETAILTableAdapter
        'Dim dtOrderDetail As MyOrderDS.ORDER_DETAILDataTable = aptOrderDetail.GetOrderDetailByOrderID(SO_NO)
        'If dtOrderDetail.Count = 0 Then
        '    Return False
        'End If
        Dim retBool As Boolean = False
        Dim p1 As New Change_SD_Order.Change_SD_Order()
        p1.Connection = New SAP.Connector.SAPConnection(System.Configuration.ConfigurationManager.AppSettings("SAP_PRD"))
        Dim OrderHeader As New Change_SD_Order.BAPISDH1, OrderHeaderX As New Change_SD_Order.BAPISDH1X
        Dim ItemIn As New Change_SD_Order.BAPISDITMTable, ItemInX As New Change_SD_Order.BAPISDITMXTable
        Dim PartNr As New Change_SD_Order.BAPIPARNRTable
        Dim Condition As New Change_SD_Order.BAPICONDTable, ScheLine As New Change_SD_Order.BAPISCHDLTable
        Dim ScheLineX As New Change_SD_Order.BAPISCHDLXTable, OrderText As New Change_SD_Order.BAPISDTEXTTable
        Dim sales_note As New Change_SD_Order.BAPISDTEXT, ext_note As New Change_SD_Order.BAPISDTEXT
        Dim op_note As New Change_SD_Order.BAPISDTEXT, retTable As New Change_SD_Order.BAPIRET2Table
        Dim ADDRTable As New Change_SD_Order.BAPIADDR1Table, PartnerChangeTable As New Change_SD_Order.BAPIPARNRCTable
        Dim Doc_Number As String = SO_NO
        OrderHeaderX.Updateflag = "U"
        OrderHeaderX.Version = "X"
        OrderHeader.Version = optyID
        p1.Connection.Open()
        Try
            p1.Bapi_Salesorder_Change("", "", New Change_SD_Order.BAPISDLS, OrderHeader, OrderHeaderX, Doc_Number, "", Condition, _
                      New Change_SD_Order.BAPICONDXTable, New Change_SD_Order.BAPIPAREXTable, New Change_SD_Order.BAPICUBLBTable, _
                      New Change_SD_Order.BAPICUINSTable, New Change_SD_Order.BAPICUPRTTable, New Change_SD_Order.BAPICUCFGTable, _
                      New Change_SD_Order.BAPICUREFTable, New Change_SD_Order.BAPICUVALTable, New Change_SD_Order.BAPICUVKTable, ItemIn, _
                      New Change_SD_Order.BAPISDITMXTable, New Change_SD_Order.BAPISDKEYTable, OrderText, ADDRTable, _
                      PartnerChangeTable, PartNr, retTable, ScheLine, ScheLineX)
            p1.CommitWork()
            retBool = True
        Catch ex As Exception
            MailUtil.SendEmail("ming.zhao@advantech.com.cn", "ming.zhao@advantech.com.cn", "Change_SD_Order", ex.ToString(), True)
        Finally
            p1.Connection.Close()
        End Try
        ReturnTable = retTable.ToADODataTable()
        For Each RetRow As DataRow In ReturnTable.Rows
            If RetRow.Item("Type").ToString().Equals("E") Then
                retBool = False
            End If
        Next
        Return retBool
    End Function
End Class
